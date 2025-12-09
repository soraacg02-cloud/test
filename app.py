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
st.set_page_config(page_title="PPT 重組生成器 (V7 穩定按鈕版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (V7 穩定按鈕版)")
st.caption("革新：V7 修正按鈕消失問題，並保留 OCR 識別功能。")

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

# --- 核心函數：V6/V7 OCR 增強版 ---
def extract_images_from_pdf_v6(pdf_stream, target_fig_text, debug=False):
    if not target_fig_text:
        return [], "Word 中未指定代表圖文字"
    
    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        matches = re.findall(r'(?:FIG\.?|Figure|图|圖)[\s\.]*([0-9]+[A-Za-z]*)', target_fig_text, re.IGNORECASE)
        if not matches:
            first_line = target_fig_text.split('\n')[0].strip().upper()
            fallback = re.search(r'([0-9]+[A-Z]*)', first_line)
            if fallback: matches = [fallback.group(1)]

        if not matches:
            return [], "無法識別任何圖號"

        target_numbers = sorted(list(set([m.upper() for m in matches])))
        
        page_blacklist_headers = [
            "BRIEF DESCRIPTION", "DETAILED DESCRIPTION", "具体实施方式", "實施方式", 
            "WHAT IS CLAIMED", "权利要求", "申請專利範圍",
            "ABSTRACT", "摘要", "BACKGROUND", "背景技術",
            "符号说明", "符號說明"
        ]

        found_page_indices = set()
        debug_logs = [] 

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

                is_text_page = False
                for header in page_blacklist_headers:
                    if header in page_text_all and len(clean_page_text_all) > 500:
                        is_text_page = True
                        break
                
                if is_text_page: continue

                match_found_strategy_1 = False
                # 策略 1: 行級別
                for b in blocks:
                    block_text = b[4].strip()
                    clean_block_text = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', block_text).upper()
                    for token in search_tokens:
                        if token in clean_block_text and len(clean_block_text) < 100:
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
                                if debug: debug_logs.append(f"✅ Found {token} (Text Layer) on P{i+1}")
                                break
                    if match_found_strategy_1: break
                
                if match_found_strategy_1: 
                    if found_this_fig: break
                    continue

                # 策略 2: 全頁級別 (Fallback)
                if len(clean_page_text_all) < 500:
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
                                if debug: debug_logs.append(f"✅ Found {token} (Full Page Text) on P{i+1}")
                                break
                
                if match_found_strategy_1:
                    if found_this_fig: break
                    continue

                # 策略 3: OCR 模式
                if len(clean_page_text_all) < 50:
                    try:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_data = pix.tobytes("png")
                        pil_image = Image.open(BytesIO(img_data))
                        
                        ocr_text = pytesseract.image_to_string(pil_image, lang='eng+chi_tra', config='--psm 11')
                        ocr_text_clean = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', ocr_text).upper()
                        
                        if debug and i < 5: debug_logs.append(f"👁️ OCR Scan P{i+1}: {ocr_text_clean[:50]}...")

                        for token in search_tokens:
                            if token in ocr_text_clean:
                                found_page_indices.add(i)
                                found_this_fig = True
                                if debug: debug_logs.append(f"✅ Found {token} (OCR) on P{i+1}")
                                break
                    except Exception as ocr_e:
                        if debug: debug_logs.append(f"⚠️ OCR Error on P{i+1}: {ocr_e}")

                if found_this_fig: break
        
        if debug and debug_logs:
            with st.expander(f"🔍 Debug 日誌: {target_numbers}"):
                st.text("\n".join(debug_logs))

        if not found_page_indices:
            return [], f"找不到圖號: {', '.join(target_numbers)} (已嘗試文字層與OCR搜尋)"

        output_images = []
        for page_idx in sorted(list(found_page_indices)):
            page = doc[page_idx]
            mat = fitz.Matrix(3, 3) 
            pix = page.get_pixmap(matrix=mat)
            output_images.append(pix.tobytes("png"))

        return output_images, f"成功 (共{len(output_images)}張)"

    except Exception as e:
        return [], f"PDF 解析錯誤: {str(e)}"

# --- 函數：提取專利號 等 ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    match = re.search(r'([a-zA-Z]{2,4}\d{4}[/]?\d+[a-zA-Z0-9]*|[a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
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
            "image_list": [], "image_name": "Word匯入", "raw_case_no": "",
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
                        "image_list": [], "image_name": "Word匯入", "raw_case_no": "",
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

# --- 側邊欄 (修正按鈕邏輯) ---
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

    st.divider()
    # === V7 修正重點：強制按鈕顯示，若無檔案則 disable ===
    btn_disabled = not word_files # 如果沒有 Word 檔案，就禁用按鈕
    run_btn = st.button("🔄 開始智能整合", type="primary", disabled=btn_disabled)

    if run_btn:
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
                
                pages_this_case = 1 
                if add_claim_slide:
                    c_groups = split_claims_text(case["claim_text"])
                    if not c_groups and case["claim_text"].strip(): pages_this_case += 1
                    else: pages_this_case += len(c_groups)
                
                start_page = current_ppt_page
                end_page = current_ppt_page + pages_this_case - 1
