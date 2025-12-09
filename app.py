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

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器 (V4 偵錯版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (V4 偵錯版)")
st.caption("革新：V4 核心邏輯 (放寬行字數限制 + 偵錯模式)，解決圖號被誤判為內文的問題。")

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

# --- 函數：搜尋 PDF 多張截圖 (V4: 寬鬆與偵錯版) ---
def extract_images_from_pdf_v4(pdf_stream, target_fig_text, debug=False):
    if not target_fig_text:
        return [], "Word 中未指定代表圖文字"
    
    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        # 1. 解析 Word 中的目標圖號
        matches = re.findall(r'(?:FIG\.?|Figure|图|圖)[\s\.]*([0-9]+[A-Za-z]*)', target_fig_text, re.IGNORECASE)
        # 備用：若 regex 失敗，嘗試抓第一行
        if not matches:
            first_line = target_fig_text.split('\n')[0].strip().upper()
            fallback = re.search(r'([0-9]+[A-Z]*)', first_line)
            if fallback:
                matches = [fallback.group(1)]

        if not matches:
            return [], "無法識別任何圖號"

        target_numbers = sorted(list(set([m.upper() for m in matches])))
        
        # 2. 定義「絕對文字頁」的標題 (碰到這些標題就整頁跳過)
        page_blacklist_headers = [
            "BRIEF DESCRIPTION", "DETAILED DESCRIPTION", "具体实施方式", "實施方式", 
            "WHAT IS CLAIMED", "权利要求", "申請專利範圍",
            "ABSTRACT", "摘要", "BACKGROUND", "背景技術",
            "符号说明", "符號說明"
        ]

        found_page_indices = set()
        debug_logs = [] # Debug 容器

        # 3. 遍歷每一個目標圖號
        for target_number in target_numbers:
            # 建立搜尋 Token
            search_tokens = [
                f"FIG{target_number}", 
                f"FIGURE{target_number}",
                f"图{target_number}", 
                f"圖{target_number}"
            ]
            
            found_this_fig = False

            for i, page in enumerate(doc):
                # 取得頁面文字區塊
                blocks = page.get_text("blocks")
                page_text_all = "".join([b[4] for b in blocks]).upper()

                # A. [頁級別過濾] 檢查黑名單標題
                is_text_page = False
                for header in page_blacklist_headers:
                    if header in page_text_all:
                        is_text_page = True
                        break
                
                # Debug: 顯示前幾頁的讀取狀況 (只在找第一個圖號時顯示，避免洗版)
                if debug and i < 5 and target_number == target_numbers[0]:
                    debug_logs.append(f"Page {i+1}: Text Length={len(page_text_all)}, IsTextPage={is_text_page}")
                    if len(page_text_all) < 200: 
                         debug_logs.append(f"   -> Content: {page_text_all[:100]}...")

                if is_text_page: 
                    continue

                # B. [行級別比對]
                for b in blocks:
                    block_text = b[4].strip()
                    # 正規化：去除非英數字與中文，轉大寫
                    clean_block_text = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', block_text).upper()
                    
                    for token in search_tokens:
                        if token in clean_block_text:
                            # 核心邏輯修正：放寬長度限制到 80 (原本30)
                            # 這是為了允許 "FIG. 1 Schematic View" 這種情況
                            if len(clean_block_text) < 80:
                                # 邊界檢查 (避免 FIG1 抓到 FIG10)
                                idx = clean_block_text.find(token)
                                is_exact_match = True
                                if idx != -1:
                                    after_idx = idx + len(token)
                                    if after_idx < len(clean_block_text):
                                        if clean_block_text[after_idx].isdigit():
                                            is_exact_match = False
                                
                                if is_exact_match:
                                    found_page_indices.add(i)
                                    found_this_fig = True
                                    if debug: debug_logs.append(f"✅ Found {token} on Page {i+1} (Text: {clean_block_text})")
                                    break
                    if found_this_fig: break
                if found_this_fig: break
        
        # 顯示 Debug 資訊
        if debug and debug_logs:
            with st.expander(f"🔍 Debug: 圖號 {target_numbers} 搜尋日誌"):
                st.text("\n".join(debug_logs))

        if not found_page_indices:
            # 若找不到，檢查是否整份 PDF 根本讀不到字 (掃描檔問題)
            total_text_len = sum([len(page.get_text()) for page in doc])
            if total_text_len < 100:
                return [], "PDF 似乎沒有文字層 (可能是純圖片掃描檔)"
            return [], f"找不到圖號: {', '.join(target_numbers)} (請嘗試開啟 Debug 模式檢查)"

        output_images = []
        for page_idx in sorted(list(found_page_indices)):
            page = doc[page_idx]
            mat = fitz.Matrix(3, 3) # 提高解析度
            pix = page.get_pixmap(matrix=mat)
            output_images.append(pix.tobytes("png"))

        return output_images, f"成功 (共{len(output_images)}張)"

    except Exception as e:
        return [], f"PDF 解析錯誤: {str(e)}"

# --- 函數：提取專利號 ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    match = re.search(r'([a-zA-Z]{2,4}\d{4}[/]?\d+[a-zA-Z0-9]*|[a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
    if match: return match.group(1)
    return ""

# --- 函數：提取詳細 Header 資訊 ---
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

# --- 函數：提取日期 (排序用) ---
def extract_date_for_sort(text):
    match = re.search(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})', text)
    if match: return f"{match.group(1)}{match.group(2).zfill(2)}{match.group(3).zfill(2)}"
    return "99999999"

# --- 函數：提取公司 (排序用) ---
def extract_company_for_sort(text):
    _, _, comp = extract_header_info_detail(text)
    if comp != "(未找到)": return comp
    return "ZZZ"

# --- 函數：正規化字串 ---
def normalize_string(s):
    if not s: return ""
    return re.sub(r'[^A-Z0-9]', '', s.upper())

# --- 函數：解析 Word 檔案 ---
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

# --- 輔助函數：分割 Claim ---
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
    debug_mode = st.checkbox("🐞 開啟偵錯模式 (Debug)", value=False, help="勾選後，會顯示 PDF 每一頁讀取到的文字，協助找出為什麼抓不到圖。")

    if word_files and st.button("🔄 開始智能整合", type="primary"):
        all_cases = []
        status_report_list = []
        for wf in word_files: all_cases.extend(parse_word_file(wf))
        
        pdf_file_map = {}
        if pdf_files:
            for pf in pdf_files:
                pdf_file_map[pf.name] = pf.read()

        match_count = 0
        current_ppt_page = 1 
        with st.spinner("處理中..."):
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
                page_str = f"P{start_page}" if start_page == end_page else f"P{start_page}-P{end_page}"
                current_ppt_page += pages_this_case

                status = {
                    "來源": case["source_file"], 
                    "案號(公開號)": case["clean_number"],
                    "公司": case["clean_company"],
                    "日期(優先權日)": case["clean_date"],
                    "對應PPT的頁碼": page_str,
                    "狀態": "未處理", "原因": "", "缺漏": ", ".join(case["missing_fields"])
                }
                
                matched_pdf = None
                norm_case_key = normalize_string(case_key)
                
                for pdf_name, pdf_bytes in pdf_file_map.items():
                    norm_pdf_name = normalize_string(pdf_name)
                    if norm_case_key and ((norm_case_key in norm_pdf_name) or (norm_pdf_name in norm_case_key)):
                        if len(norm_case_key) > 5:
                            matched_pdf = pdf_bytes
                            break
                
                if matched_pdf:
                    # 使用 V4 函數 (含 Debug)
                    img_list, msg = extract_images_from_pdf_v4(matched_pdf, target_fig, debug=debug_mode)
                    if img_list:
                        case["image_list"] = img_list
                        status["狀態"] = f"✅ 成功 ({len(img_list)}張)"
                        match_count += 1
                    else:
                        status["狀態"] = "⚠️ 缺圖"; status["原因"] = msg
                else:
                    if not target_fig: status["狀態"] = "⚠️ 缺資訊"; status["原因"] = "Word無代表圖"
                    else: status["狀態"] = "❌ 無PDF"; status["原因"] = f"找不到PDF: {case_key}"
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
                    st.image(data['image_list'][0], caption=f"共 {len(data['image_list'])} 張圖", use_column_width=True)
                else:
                    st.warning("無圖片")
                full_claim_text = data['claim_text']
                claims_preview = split_claims_text(full_claim_text)
                count_claims = len(claims_preview) if full_claim_text else 0
                st.caption(f"Claim: {count_claims} 組")

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data, need_claim_slide):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        for data in slides_data:
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
                    
                    left, top, width, height = Inches(0.5), Inches(2.5), Inches(12.3), Inches(4.5)
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
        cols = ["來源", "案號(公開號)", "公司", "日期(優先權日)", "對應PPT的頁碼", "狀態", "原因", "缺漏"]
        st.dataframe(df[cols], hide_index=True)
