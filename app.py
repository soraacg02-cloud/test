import streamlit as st
import streamlit.components.v1 as components
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE 
from io import BytesIO
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
import fitz  # PyMuPDF
import re
import pandas as pd
from PIL import Image
import pytesseract

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器 (V17 寬鬆比對版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (V17 寬鬆比對版)")
st.caption("更新：V17 加入「核心數字比對」機制。當 PDF 檔名包含額外的零 (如 us000123...) 時，程式能透過識別關鍵數字串 (如 123) 成功找到對應檔案。")

# === NBLM 提示詞區塊 ===
nblm_prompt = """根據上傳的所有來源，分開整理出以下重點(不要表格)：

1. 案號 / 日期 / 公司： *(案號依據"公開號"、日期依據"優先權日"、公司依據"申請人")
2. 解決問題：
3. 發明精神：*(不要有公式)
4. 一句重點： *(用來描述發明特徵重點，20字)
5. 代表圖：*(根據發明精神建議3張最可以說明發明精神的圖片，範例:FIG.3)
6. 獨立項claim： *(分組且分行條列式+對應的代表圖，claim要(1)有位階縮排 (2)claim的元件要有標號 (3)對應的claim號碼)"""

st.info("💡 **NBLM 使用提示詞** (點擊下方綠色按鈕一鍵複製)")

components.html(
    f"""
    <html>
    <head><meta charset="utf-8"></head>
    <body style="font-family: sans-serif; margin: 0; padding: 0;">
        <div style="display: flex; flex-direction: column; align-items: flex-start;">
            <textarea id="copyTarget" style="opacity: 0; position: absolute; z-index: -1;">{nblm_prompt}</textarea>
            <div style="background-color: #f0f2f6; padding: 15px; border-radius: 10px; white-space: pre-wrap; font-size: 14px; color: #31333F; border: 1px solid #d6d6d6; width: 95%; margin-bottom: 10px;">{nblm_prompt}</div>
            <button onclick="copyFunction()" style="background-color: #00CC66; color: white; border: none; padding: 12px 24px; font-size: 16px; font-weight: bold; border-radius: 8px; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">📋 點我一鍵複製提示詞</button>
            <span id="statusParams" style="color: #00CC66; font-weight: bold; margin-left: 10px; opacity: 0; transition: opacity 0.5s;">✅ 複製成功！</span>
        </div>
        <script>
        function copyFunction() {{
            var copyText = document.getElementById("copyTarget");
            copyText.select();
            navigator.clipboard.writeText(copyText.value).then(function() {{
                var status = document.getElementById("statusParams");
                status.style.opacity = '1';
                setTimeout(function(){{ status.style.opacity = '0'; }}, 2000);
            }});
        }}
        </script>
    </body>
    </html>
    """,
    height=360
)
st.divider()

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []
if 'status_report' not in st.session_state:
    st.session_state['status_report'] = []
if 'debug_logs_map' not in st.session_state:
    st.session_state['debug_logs_map'] = {}

# --- 輔助函數：遍歷 Word ---
def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    else:
        raise ValueError("只支援讀取整份 Document")
    for child in parent_elm.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('tbl'):
            yield Table(child, parent)

# --- 核心函數：V13 動態平衡版邏輯 ---
def extract_images_from_pdf_v13(pdf_stream, target_fig_text, case_key, debug=False, log_prefix=""):
    if not target_fig_text:
        return [], f"{log_prefix}未指定圖號"
    
    try:
        if hasattr(pdf_stream, 'seek'):
            pdf_stream.seek(0)
            
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        matches = re.findall(r'(?:FIG\.?|Figure|图|圖)[\s\.]*([0-9]+[A-Za-z]*)', target_fig_text, re.IGNORECASE)
        if not matches:
            first_line = target_fig_text.split('\n')[0].strip().upper()
            fallback = re.search(r'([0-9]+[A-Z]*)', first_line)
            if fallback: matches = [fallback.group(1)]

        if not matches:
            return [], f"{log_prefix}無法識別任何圖號"

        target_numbers = sorted(list(set([m.upper() for m in matches])))
        
        # V13 參數
        PAGE_TEXT_THRESHOLD_OCR = 800  
        PAGE_TEXT_THRESHOLD_RAW = 600 
        LONG_SENTENCE_LIMIT = 80 
        MAX_LONG_SENTENCES = 3
        LINE_LENGTH_LIMIT = 30

        page_blacklist_headers = [
            "BRIEF DESCRIPTION", "DETAILED DESCRIPTION", "具体实施方式", "實施方式", 
            "WHAT IS CLAIMED", "权利要求", "申請專利範圍", "圖式簡單說明", "【圖式簡單說明】",
            "ABSTRACT", "摘要", "BACKGROUND", "背景技術",
            "符号说明", "符號說明"
        ]

        SENTENCE_STOPWORDS = ["為", "係", "所示", "關於", "參照", "參考", "EXAMPLE", "EMBODIMENT", "SHOWS", "REFER"]

        found_page_indices = set()
        debug_logs = [] 
        debug_logs.append(f"{log_prefix}🎯 目標: {target_numbers}")

        for target_number in target_numbers:
            search_tokens = [
                f"FIG{target_number}", f"FIGURE{target_number}",
                f"图{target_number}", f"圖{target_number}"
            ]
            
            found_this_fig = False

            for i, page in enumerate(doc):
                blocks = page.get_text("blocks")
                page_text_all = "".join([b[4] for b in blocks]).upper()
                clean_page_text_all = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', page_text_all)
                page_text_len = len(clean_page_text_all)

                is_blacklist_page = False
                for header in page_blacklist_headers:
                    if header in page_text_all:
                        is_blacklist_page = True
                        if debug and i < 15: debug_logs.append(f"{log_prefix}🚫 Skip P{i+1} (Header: {header})")
                        break
                if is_blacklist_page: continue

                long_sentence_count = 0
                for b in blocks:
                    if len(re.sub(r'\s+', '', b[4])) > LONG_SENTENCE_LIMIT:
                        long_sentence_count += 1
                
                if long_sentence_count > MAX_LONG_SENTENCES:
                    if debug and i < 15: debug_logs.append(f"{log_prefix}🚫 Skip P{i+1} (Raw: Long sentences)")
                    continue

                if page_text_len > PAGE_TEXT_THRESHOLD_RAW:
                    if debug and i < 15: debug_logs.append(f"{log_prefix}🚫 Skip P{i+1} (Raw Heavy: {page_text_len})")
                    continue

                match_found_strategy_1 = False
                for b in blocks:
                    block_text = b[4].strip()
                    clean_block_text = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', block_text).upper()
                    
                    for token in search_tokens:
                        if token in clean_block_text:
                            if len(clean_block_text) > LINE_LENGTH_LIMIT: continue
                            
                            is_sentence = False
                            for stopword in SENTENCE_STOPWORDS:
                                if stopword in clean_block_text:
                                    is_sentence = True
                                    break
                            if is_sentence: continue 

                            idx = clean_block_text.find(token)
                            is_exact_match = True
                            if idx != -1:
                                after_idx = idx + len(token)
                                if after_idx < len(clean_block_text) and clean_block_text[after_idx].isdigit():
                                    is_exact_match = False
                            
                            if is_exact_match:
                                found_page_indices.add(i)
                                found_this_fig = True
                                match_found_strategy_1 = True
                                if debug: debug_logs.append(f"{log_prefix}✅ Found {token} (Text Layer) on P{i+1}")
                                break
                    if match_found_strategy_1: break
                
                if match_found_strategy_1: 
                    if found_this_fig: break
                    continue

                if page_text_len < PAGE_TEXT_THRESHOLD_RAW:
                    for token in search_tokens:
                        if token in clean_page_text_all:
                            idx = clean_page_text_all.find(token)
                            is_exact_match = True
                            if idx != -1:
                                after_idx = idx + len(token)
                                if after_idx < len(clean_page_text_all) and clean_page_text_all[after_idx].isdigit():
                                    is_exact_match = False
                            if is_exact_match:
                                found_page_indices.add(i)
                                found_this_fig = True
                                match_found_strategy_1 = True
                                if debug: debug_logs.append(f"{log_prefix}✅ Found {token} (Full Page) on P{i+1}")
                                break
                
                if match_found_strategy_1:
                    if found_this_fig: break
                    continue

                if page_text_len < 200: 
                    try:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_data = pix.tobytes("png")
                        pil_image = Image.open(BytesIO(img_data))
                        
                        ocr_text = pytesseract.image_to_string(pil_image, lang='eng+chi_tra', config='--psm 11')
                        ocr_text_clean = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', ocr_text).upper()
                        ocr_len = len(ocr_text_clean)

                        if debug and i < 15: debug_logs.append(f"{log_prefix}👁️ OCR P{i+1} Len: {ocr_len}")

                        if ocr_len > PAGE_TEXT_THRESHOLD_OCR:
                            if debug: debug_logs.append(f"{log_prefix}   -> Skip P{i+1} (OCR Heavy)")
                            continue
                        
                        ocr_lines = ocr_text.split('\n')
                        long_sentence_count_ocr = 0
                        for line in ocr_lines:
                             clean_line_len = len(re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', line))
                             if clean_line_len > LONG_SENTENCE_LIMIT:
                                 long_sentence_count_ocr += 1
                        
                        if long_sentence_count_ocr > MAX_LONG_SENTENCES:
                             if debug: debug_logs.append(f"{log_prefix}   -> Skip P{i+1} (OCR Long Sentences)")
                             continue

                        for line in ocr_lines:
                            clean_line = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', line).upper()
                            
                            for token in search_tokens:
                                if token in clean_line:
                                    if len(clean_line) > LINE_LENGTH_LIMIT: continue
                                    
                                    is_sentence_ocr = False
                                    for stopword in SENTENCE_STOPWORDS:
                                        if stopword in clean_line:
                                            is_sentence_ocr = True
                                            break
                                    if is_sentence_ocr: continue

                                    found_page_indices.add(i)
                                    found_this_fig = True
                                    if debug: debug_logs.append(f"{log_prefix}✅ Found {token} (OCR) on P{i+1}")
                                    break
                            if found_this_fig: break

                    except Exception as ocr_e:
                        if debug: debug_logs.append(f"{log_prefix}⚠️ OCR Error on P{i+1}: {ocr_e}")

                if found_this_fig: break
        
        if debug:
            if case_key not in st.session_state['debug_logs_map']:
                st.session_state['debug_logs_map'][case_key] = ""
            st.session_state['debug_logs_map'][case_key] += "\n".join(debug_logs) + "\n\n"

        if not found_page_indices:
            return [], f"{log_prefix}找不到圖號: {', '.join(target_numbers)}"

        output_images = []
        for page_idx in sorted(list(found_page_indices)):
            page = doc[page_idx]
            mat = fitz.Matrix(4, 4) 
            pix = page.get_pixmap(matrix=mat)
            output_images.append(pix.tobytes("png"))

        return output_images, f"成功 ({len(output_images)}張)"

    except Exception as e:
        return [], f"{log_prefix}PDF 解析錯誤: {str(e)}"

# --- 函數：提取專利號 (V16 修正：去除逗號) ---
def extract_patent_number_from_text(text):
    if "：" in text: text = text.replace("：", ":")
    if ":" in text:
        content = text.split(":", 1)[1]
    else:
        content = text

    clean_text = content.replace(" ", "").replace(",", "").strip().upper()
    match = re.search(r'([A-Z]{2,4}\d{4,}[A-Z0-9]*)', clean_text)
    if match: return match.group(1)
    
    return ""

def extract_header_info_detail(raw_text):
    number = "(未找到)"
    date = "(未找到)"
    company = "(未找到)"
    
    extracted_no = extract_patent_number_from_text(raw_text)
    if extracted_no: number = extracted_no
    else:
        match_no = re.search(r'(?:公開號|案號)[:：\s]*([^\n]+)', raw_text)
        if match_no:
            raw_no = match_no.group(1)
            raw_no = re.split(r'\s+(?:日期|公司|申請人)[:：]', raw_no)[0]
            number = raw_no.strip()

    match_date = re.search(r'(?:日期)[:：\s]*(\d{4}[./-]\d{1,2}[./-]\d{1,2})', raw_text)
    if match_date: date = match_date.group(1).strip()
    else:
        match_date_backup = re.search(r'(\d{4}[./-]\d{1,2}[./-]\d{1,2})', raw_text)
        if match_date_backup: date = match_date_backup.group(1).strip()

    matches = re.findall(r'(?:公司|申請人)[:：\s]*(.*?)(?=\s+(?:公開號|案號|日期)[:：]|$)', raw_text)
    if matches:
        for candidate in reversed(matches):
            clean_cand = candidate.strip()
            if len(clean_cand) > 1 and "公開號" not in clean_cand:
                company = clean_cand
                break

    return number, date, company

def extract_date_for_sort(text):
    match = re.search(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})', text)
    if match: return f"{match.group(1)}{match.group(2).zfill(2)}{match.group(3).zfill(2)}"
    return "99999999"

def extract_company_for_sort(text):
    _, _, comp = extract_header_info_detail(text)
    if comp != "(未找到)": return comp
    return "ZZZ"

def normalize_string(s):
    if not s: return ""
    return re.sub(r'[^A-Z0-9]', '', s.upper())

def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        current_case = {
            "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "", "claim_text": "",
            "image_list": [], "claim_image_list": [], "image_name": "Word匯入", "raw_case_no": "",
            "clean_number": "", "clean_date": "", "clean_company": "", 
            "sort_date": "99999999", "sort_company": "ZZZ",
            "source_file": uploaded_docx.name, "missing_fields": []
        }
        current_field = None 
        
        all_lines = []
        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                if block.text.strip(): all_lines.append(block.text.strip())
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if p.text.strip(): all_lines.append(p.text.strip())
        
        for text in all_lines:
            if "案號" in text or "索號" in text:
                if current_case["case_info"] and current_field != "case_info_block":
                    nb, dt, cp = extract_header_info_detail(current_case["case_info"])
                    current_case["clean_number"] = nb
                    current_case["clean_date"] = dt
                    current_case["clean_company"] = cp
                    if not current_case["problem"]: current_case["missing_fields"].append("解決問題")
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "", "claim_text": "",
                        "image_list": [], "claim_image_list": [], "image_name": "Word匯入", "raw_case_no": "",
                        "clean_number": "", "clean_date": "", "clean_company": "",
                        "sort_date": "99999999", "sort_company": "ZZZ",
                        "source_file": uploaded_docx.name, "missing_fields": []
                    }
                current_field = "case_info_block"
                current_case["case_info"] = text
                
                nb, dt, cp = extract_header_info_detail(text)
                if dt != "(未找到)": current_case["sort_date"] = dt.replace(".", "").replace("/", "").replace("-", "")
                if cp != "(未找到)": current_case["sort_company"] = cp
                if nb != "(未找到)": current_case["raw_case_no"] = nb
                continue

            if "解決問題" in text:
                current_field = "problem"
                current_case["problem"] = re.sub(r'^[0-9.．]*\s*解決問題[:：]?\s*', '', text)
                continue
            elif "發明精神" in text:
                current_field = "spirit"
                current_case["spirit"] = re.sub(r'^[0-9.．]*\s*發明精神[:：]?\s*', '', text)
                continue
            elif "重點" in text:
                current_field = "key_point"
                current_case["key_point"] = re.sub(r'^[0-9.．]*\s*(一句)?重點[:：]?\s*', '', text)
                continue
            elif "代表圖" in text:
                current_field = "rep_fig"
                current_case["rep_fig_text"] = re.sub(r'^[0-9.．]*\s*代表圖[:：]?\s*', '', text).strip()
                continue
            elif "獨立項" in text or ("claim" in text.lower() and "6" in text):
                current_field = "claim"
                content = re.sub(r'^[0-9.．]*\s*(獨立項)?(claim)?[:：]?\s*', '', text, flags=re.IGNORECASE).strip()
                current_case["claim_text"] = content
                continue

            if current_field == "case_info_block":
                current_case["case_info"] += "\n" + text
                nb, dt, cp = extract_header_info_detail(current_case["case_info"])
                if dt != "(未找到)": current_case["sort_date"] = dt.replace(".", "").replace("/", "").replace("-", "")
                if cp != "(未找到)": current_case["sort_company"] = cp
                if nb != "(未找到)": current_case["raw_case_no"] = nb
            elif current_field == "rep_fig":
                current_case["rep_fig_text"] += "\n" + text
            elif current_field == "problem":
                current_case["problem"] += "\n" + text
            elif current_field == "spirit":
                current_case["spirit"] += "\n" + text
            elif current_field == "key_point":
                current_case["key_point"] += "\n" + text
            elif current_field == "claim": 
                current_case["claim_text"] += "\n" + text

        if current_case["case_info"]:
            nb, dt, cp = extract_header_info_detail(current_case["case_info"])
            current_case["clean_number"] = nb
            current_case["clean_date"] = dt
            current_case["clean_company"] = cp
            if not current_case["problem"]: current_case["missing_fields"].append("解決問題")
            cases.append(current_case)
        return cases
    except Exception as e:
        st.error(f"解析 Word 錯誤 ({uploaded_docx.name}): {e}")
        return []

def split_claims_text(full_text):
    if not full_text: return []
    lines = full_text.split('\n')
    claims = []
    current_chunk = []
    header_pattern = re.compile(r'(\(Claim\s*\d+\)|^\s*(Claim|獨立項)\s*\d+|^\s*\d+\.\s)', re.IGNORECASE)
    for line in lines:
        if header_pattern.search(line):
            if current_chunk:
                if "".join(current_chunk).strip(): claims.append(current_chunk)
            current_chunk = [line]
        else:
            current_chunk.append(line)
    if current_chunk and "".join(current_chunk).strip(): claims.append(current_chunk)
    return claims

def parse_fig_number_from_claim(claim_text):
    if not claim_text: return None
    matches = re.findall(r'(?:FIG\.?|Figure|图|圖)[\s\.]*([0-9]+[A-Za-z]*)', claim_text, re.IGNORECASE)
    if matches:
        return "FIG. " + ", FIG. ".join(sorted(list(set(matches))))
    return None

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料")
    word_files = st.file_uploader("Word 檔案 (可多選)", type=['docx'], accept_multiple_files=True)
    pdf_files = st.file_uploader("PDF 檔案 (可多選)", type=['pdf'], accept_multiple_files=True)
    st.divider()
    st.header("2. 輸出設定")
    add_claim_slide = st.checkbox("✅ 是否產生 Claim 分頁", value=False, help="勾選後，程式會自動識別獨立項數量，並為每一組獨立項產生一頁")
    
    st.divider()
    st.header("3. 進階除錯")
    debug_mode = st.checkbox("🐞 開啟偵錯模式 (Debug)", value=False, help="勾選後，會顯示詳細的識別日誌，包含 OCR 的辨識結果。")
    
    if debug_mode and st.session_state['debug_logs_map']:
        st.caption("📜 歷史 Debug 紀錄 (點擊展開)")
        for key, log in st.session_state['debug_logs_map'].items():
            with st.expander(f"Case: {key}"):
                st.text(log)

    st.divider()
    run_btn = st.button("🔄 開始智能整合", type="primary")

    if run_btn:
        st.session_state['debug_logs_map'] = {}
        
        if not word_files:
            st.warning("⚠️ 請先上傳 Word 檔案！")
            st.stop()

        all_cases = []
        status_report_list = []
        for wf in word_files: all_cases.extend(parse_word_file(wf))
        
        pdf_file_map = {}
        if pdf_files:
            for pf in pdf_files:
                pdf_file_map[pf.name] = pf.read()

        match_count = 0
        current_ppt_page = 1 
        with st.spinner("處理中... (若啟動 OCR 可能需要較長時間，請耐心等候)"):
            all_cases.sort(key=lambda x: (x["sort_company"].upper(), x["sort_date"]))
            for case in all_cases:
                case_key = case["raw_case_no"]
                target_fig = case["rep_fig_text"]
                claim_text_content = case["claim_text"]
                
                pages_this_case = 1 
                if add_claim_slide:
                    c_groups = split_claims_text(claim_text_content)
                    if not c_groups and claim_text_content.strip(): pages_this_case += 1
                    else: pages_this_case += len(c_groups)
                
                start_page = current_ppt_page
                end_page = current_ppt_page + pages_this_case - 1
                page_str = f"P{start_page}" if start_page == end_page else f"P{start_page}-P{end_page}"
                current_ppt_page += pages_this_case

                status = {
                    "來源": case["source_file"], 
                    "案號(公開號)": case["clean_number"],
                    "公司": case["clean_company"],
                    "日期(優先權日)": case["clean_date"],
                    "對應PPT的頁碼": page_str,
                    "狀態": "未處理", "原因": "", "缺漏": ", ".join(case["missing_fields"]),
                    "Claim圖狀態": "N/A", "Claim圖說明": ""
                }
                
                matched_pdf = None
                norm_case_key = normalize_string(case_key)
                
                # === V17 修正：寬鬆比對邏輯 (Smart Matching) ===
                for pdf_name, pdf_bytes in pdf_file_map.items():
                    norm_pdf_name = normalize_string(pdf_name)
                    
                    # 1. 精準比對 (舊邏輯)
                    if norm_case_key and ((norm_case_key in norm_pdf_name) or (norm_pdf_name in norm_case_key)):
                        if len(norm_case_key) > 5:
                            matched_pdf = pdf_bytes
                            break
                    
                    # 2. 核心數字比對 (新邏輯)
                    # 提取 Case 中的純數字: US11226533B2 -> 11226533
                    case_digits = re.sub(r'\D', '', case_key)
                    if len(case_digits) >= 4 and case_digits in norm_pdf_name:
                        matched_pdf = pdf_bytes
                        break
                
                if matched_pdf:
                    # 1. 抓取主要代表圖
                    img_list_main, msg_main = extract_images_from_pdf_v13(matched_pdf, target_fig, case_key, debug=debug_mode, log_prefix="[Main] ")
                    
                    if img_list_main:
                        case["image_list"] = img_list_main
                        status["狀態"] = f"✅ 成功 ({len(img_list_main)}張)"
                        match_count += 1
                    else:
                        status["狀態"] = "⚠️ 缺圖"; status["原因"] = msg_main

                    # 2. 抓取 Claim 附圖
                    if add_claim_slide:
                        specific_claim_fig = parse_fig_number_from_claim(claim_text_content)
                        img_list_claim = []
                        msg_claim = ""
                        
                        if specific_claim_fig:
                            img_list_claim, msg_claim = extract_images_from_pdf_v13(matched_pdf, specific_claim_fig, case_key, debug=debug_mode, log_prefix="[Claim] ")
                            if img_list_claim:
                                status["Claim圖狀態"] = f"✅ 專屬 ({len(img_list_claim)}張)"
                                status["Claim圖說明"] = f"找到指定圖: {specific_claim_fig}"
                            else:
                                if img_list_main:
                                    img_list_claim = img_list_main
                                    status["Claim圖狀態"] = "⚠️ 沿用主圖"
                                    status["Claim圖說明"] = f"指定圖 ({specific_claim_fig}) 抓取失敗: {msg_claim}"
                                else:
                                    status["Claim圖狀態"] = "❌ 缺圖"
                                    status["Claim圖說明"] = "指定失敗且無主圖"
                        else:
                            if img_list_main:
                                img_list_claim = img_list_main
                                status["Claim圖狀態"] = "✅ 同主圖"
                                status["Claim圖說明"] = "未指定"
                            else:
                                status["Claim圖狀態"] = "❌ 缺圖"
                                status["Claim圖說明"] = "未指定且無主圖"
                        
                        case["claim_image_list"] = img_list_claim

                else:
                    if not target_fig: status["狀態"] = "⚠️ 缺資訊"; status["原因"] = "Word無代表圖"
                    else: status["狀態"] = "❌ 無PDF"; status["原因"] = f"找不到PDF: {case_key} (已嘗試寬鬆比對)"
                status_report_list.append(status)

        if all_cases:
            st.session_state['slides_data'] = all_cases
            st.session_state['status_report'] = status_report_list
            st.success(f"完成！共 {len(all_cases)} 筆資料。")
        else:
            st.warning("無資料。")

    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除重來"):
            st.session_state['slides_data'] = []
            st.session_state['status_report'] = []
            st.session_state['debug_logs_map'] = {}
            st.rerun()

# --- 主畫面 ---
if not st.session_state['slides_data']:
    st.info("👈 請先上傳檔案。")
else:
    st.subheader(f"📋 預覽 (已排序: 申請人 -> 日期)")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**Case {i+1}**")
                st.caption(f"{data['clean_company']} | {data['clean_date']}")
                st.text(f"{data['clean_number']}")
                if data['image_list']:
                    st.image(data['image_list'][0], caption=f"主圖 ({len(data['image_list'])})", use_column_width=True)
                
                if data.get('claim_image_list'):
                     st.image(data['claim_image_list'][0], caption=f"Claim 用圖 ({len(data['claim_image_list'])})", use_column_width=True)
                
                full_claim_text = data['claim_text']
                claims_preview = split_claims_text(full_claim_text)
                count_claims = len(claims_preview) if full_claim_text else 0
                st.caption(f"Claim: {count_claims} 組")

    def generate_ppt(slides_data, need_claim_slide):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        for data in slides_data:
            # === Main Slide ===
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(2.0)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame; tf.word_wrap = True
            p1 = tf.add_paragraph(); p1.text = f"公開號：{data['clean_number']}"; p1.font.size = Pt(20); p1.font.bold = True
            p2 = tf.add_paragraph(); p2.text = f"日期：{data['clean_date']}"; p2.font.size = Pt(20); p2.font.bold = True
            p3 = tf.add_paragraph(); p3.text = f"公司：{data['clean_company']}"; p3.font.size = Pt(20); p3.font.bold = True

            img_left = Inches(5.5); img_top = Inches(0.5); img_width = Inches(7.0)
            img_list = data.get('image_list', [])
            
            if img_list:
                num_imgs = len(img_list)
                img_w = (7.0 / num_imgs) - 0.1
                img_h = 3.0
                for idx, img_bytes in enumerate(img_list):
                    this_left = 5.5 + (idx * (img_w + 0.1))
                    slide.shapes.add_picture(BytesIO(img_bytes), Inches(this_left), Inches(0.5), height=Inches(img_h))
                
                text_top = Inches(3.6)
                text_height = Inches(1.0)
                txBox = slide.shapes.add_textbox(img_left, text_top, img_width, text_height)
                tf = txBox.text_frame; tf.word_wrap = True
                content = data['rep_fig_text'] if data['rep_fig_text'].strip() else ""
                for line in content.split('\n'):
                    if line.strip():
                        p = tf.add_paragraph(); p.text = line.strip(); p.font.size = Pt(14)
            else:
                img_height = Inches(4.0)
                txBox = slide.shapes.add_textbox(img_left, img_top, img_width, img_height)
                tf = txBox.text_frame; tf.word_wrap = True
                content = data['rep_fig_text'] if data['rep_fig_text'].strip() else "無代表圖資訊"
                for line in content.split('\n'):
                    if line.strip():
                        p = tf.add_paragraph(); p.text = line.strip(); p.font.size = Pt(16)

            left, top, width, height = Inches(0.5), Inches(4.8), Inches(12.3), Inches(1.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame; tf.word_wrap = True
            p1 = tf.add_paragraph(); p1.text = "• 解決問題：" + data['problem']; p1.font.size = Pt(18); p1.space_after = Pt(12)
            p2 = tf.add_paragraph(); p2.text = "• 發明精神：" + data['spirit']; p2.font.size = Pt(18)

            left, top, width, height = Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.8)
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
            shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(255, 192, 0); shape.line.color.rgb = RGBColor(255, 192, 0)
            p = shape.text_frame.paragraphs[0]; p.text = data['key_point']; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(20); p.font.bold = True
            shape.text_frame.vertical_anchor = MSO_SHAPE.RECTANGLE

            # === Claim Slides ===
            if need_claim_slide:
                claims_groups = split_claims_text(data['claim_text'])
                if not claims_groups and data['claim_text'].strip():
                      claims_groups = [data['claim_text'].split('\n')]

                for claim_lines in claims_groups:
                    slide_c = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(2.0)
                    txBox = slide_c.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame; tf.word_wrap = True
                    p1 = tf.add_paragraph(); p1.text = f"公開號：{data['clean_number']}"; p1.font.size = Pt(20); p1.font.bold = True
                    p2 = tf.add_paragraph(); p2.text = f"日期：{data['clean_date']}"; p2.font.size = Pt(20); p2.font.bold = True
                    p3 = tf.add_paragraph(); p3.text = f"公司：{data['clean_company']}"; p3.font.size = Pt(20); p3.font.bold = True
                    
                    # 貼上 Claim 圖片 (如果有)
                    claim_imgs = data.get('claim_image_list', [])
                    if claim_imgs:
                        img_left = Inches(5.5); img_top = Inches(0.5)
                        num_imgs = len(claim_imgs)
                        img_w = (7.0 / num_imgs) - 0.1
                        img_h = 3.0
                        for idx, img_bytes in enumerate(claim_imgs):
                            this_left = 5.5 + (idx * (img_w + 0.1))
                            slide_c.shapes.add_picture(BytesIO(img_bytes), Inches(this_left), Inches(0.5), height=Inches(img_h))

                    # 根據有無圖片調整文字框位置
                    left, width = Inches(0.5), Inches(12.3)
                    if claim_imgs:
                         top = Inches(3.6); height = Inches(3.4)
                    else:
                         top = Inches(2.5); height = Inches(4.5)

                    txBox = slide_c.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame; tf.word_wrap = True
                    
                    p_title = tf.add_paragraph()
                    p_title.text = "【獨立項 Claim】"
                    p_title.font.size = Pt(24); p_title.font.bold = True; p_title.font.color.rgb = RGBColor(0, 112, 192)
                    p_title.space_after = Pt(10)
                    
                    for line in claim_lines:
                        clean_line = line.strip()
                        if clean_line:
                            p = tf.add_paragraph()
                            p.text = clean_line
                            p.font.size = Pt(14) 
                            p.space_after = Pt(4)
                            
                            if line.startswith('\t') or line.startswith('    '):
                                p.level = 1
                            elif clean_line.startswith(('o ', '○', '-', '•', '●')):
                                p.level = 1
                            elif clean_line.startswith(('▪', '■')):
                                p.level = 2
                            elif re.match(r'^(\(\d+\)|\d+\.|\d+\))', clean_line):
                                if "Claim" in clean_line or "獨立項" in clean_line:
                                    p.level = 0
                                    p.font.bold = True
                                else:
                                    p.level = 1

        return prs

    st.divider()
    if st.button("🚀 生成 PowerPoint (.pptx)", type="primary"):
        prs = generate_ppt(st.session_state['slides_data'], add_claim_slide)
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        st.download_button("📥 下載 PPT", binary_output, "slides_with_claims.pptx")

    st.divider()
    st.subheader("📊 診斷報告")
    if st.session_state['status_report']:
        df = pd.DataFrame(st.session_state['status_report'])
        cols = ["來源", "案號(公開號)", "公司", "日期(優先權日)", "對應PPT的頁碼", "狀態", "原因", "Claim圖狀態", "Claim圖說明", "缺漏"]
        st.dataframe(df[cols], hide_index=True)
