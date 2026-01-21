import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import xlsxwriter
import io
import json
import re
import os
import time
from PIL import Image

# -----------------------------------------------------------
# [1] 페이지 설정
# -----------------------------------------------------------
st.set_page_config(page_title="논문 분석 Pro", layout="wide")

# -----------------------------------------------------------
# [2] 메인 UI
# -----------------------------------------------------------
# 버전 업데이트: 5.8 -> 5.9
st.title("📑 논문 분석 Pro [ver5.9]")
st.caption("✅ 이미지 내 텍스트(Fig/Table) 분석 | 구조 분석 보완")

# -----------------------------------------------------------
# [3] 사이드바
# -----------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 설정")
    default_key = ''
    api_key_input = st.text_input("Google API Key", value=default_key, type="password")

    if not api_key_input:
        st.warning("👈 API 키를 입력해주세요.")
        st.stop()

    genai.configure(api_key=api_key_input, transport='rest')

    st.subheader("🤖 AI 모델 선택")
    try:
        available_models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.replace('models/', '')
                available_models.append(name)

        preferred = ['gemini-1.5-flash', 'gemini-2.5-flash']
        available_models.sort(key=lambda x: (x not in preferred, x))

        if not available_models:
            st.error("사용 가능한 모델이 없습니다.")
            st.stop()

        selected_model_name = st.selectbox(
            "✅ 사용 가능한 모델 목록",
            available_models,
            index=0
        )
        SELECTED_MODEL_NAME = f"models/{selected_model_name}"
        st.success(f"연결됨: {selected_model_name}")

    except Exception as e:
        st.error(f"모델 목록 오류: {e}")
        st.stop()

    # [수정됨] 불필요한 '이미지 정밀 판독' 옵션 및 관련 UI 제거

model = genai.GenerativeModel(SELECTED_MODEL_NAME)


# [수정됨] vision_model 제거 (더 이상 사용하지 않음)


# -----------------------------------------------------------
# [4] 유틸리티 함수
# -----------------------------------------------------------

def normalize_id(ref_text):
    """이미지 ID 정규화"""
    nums = re.findall(r'\d+', str(ref_text))
    return f"Image_{nums[0]}" if nums else None


def merge_nearby_rectangles(rects, distance=20):
    """사각형 병합 (스마트 머지)"""
    if not rects: return []
    rects.sort(key=lambda r: (r.y0, r.x0))
    merged = []
    while rects:
        current = rects.pop(0)
        has_merged = True
        while has_merged:
            has_merged = False
            rest = []
            for r in rects:
                expanded_current = fitz.Rect(current.x0 - distance, current.y0 - distance,
                                             current.x1 + distance, current.y1 + distance)
                if expanded_current.intersects(r):
                    current = current | r
                    has_merged = True
                else:
                    rest.append(r)
            rects = rest
        merged.append(current)
    return merged


# -----------------------------------------------------------
# [5] 핵심 로직 함수
# -----------------------------------------------------------

def extract_data_from_pdf(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    final_text_content = ""
    image_counter = 1

    all_captions = []
    all_images_info = []

    # 1. 정보 수집
    for page_num, page in enumerate(doc):
        text_blocks = page.get_text("blocks")
        for b in text_blocks:
            text = b[4].strip()
            # 캡션 후보 식별
            if (text.startswith("Fig") or text.startswith("Table")) and len(text) < 500:
                bbox = fitz.Rect(b[0], b[1], b[2], b[3])
                cap_type = "Figure" if text.startswith("Fig") else "Table"
                label_match = re.match(r"(Fig\.?|Table)\s*\d+", text)
                label = label_match.group(0) if label_match else cap_type

                all_captions.append({
                    "page": page_num, "bbox": bbox, "text": text,
                    "type": cap_type, "label": label, "matched_img_id": None
                })

        image_list = page.get_images(full=True)
        raw_rects = []
        for img in image_list:
            xref = img[0]
            img_rects = page.get_image_rects(xref)
            for r in img_rects:
                if r.width < 10 or r.height < 10: continue
                raw_rects.append(r)

        merged_rects = merge_nearby_rectangles(raw_rects, distance=20)

        for rect in merged_rects:
            img_id = f"Image_{image_counter}"
            all_images_info.append({
                "id": img_id, "page": page_num, "bbox": rect, "matched_caption": None
            })
            image_counter += 1

    # 2. 위치 기반 매칭 (보조)
    for cap in all_captions:
        best_img = None
        min_score = float('inf')
        candidates = [img for img in all_images_info if img["page"] == cap["page"] and img["matched_caption"] is None]

        for img in candidates:
            # 방향 규칙
            if cap["type"] == "Figure" and cap["bbox"].y0 < img["bbox"].y1: continue
            if cap["type"] == "Table" and cap["bbox"].y1 > img["bbox"].y0: continue

            v_dist = max(0, cap["bbox"].y0 - img["bbox"].y1) if cap["type"] == "Figure" else max(0,
                                                                                                 img["bbox"].y0 - cap[
                                                                                                     "bbox"].y1)
            cap_center_x = (cap["bbox"].x0 + cap["bbox"].x1) / 2
            img_center_x = (img["bbox"].x0 + img["bbox"].x1) / 2
            h_align_dist = abs(cap_center_x - img_center_x)

            if h_align_dist > 150: continue

            total_score = v_dist + (h_align_dist * 2.5)
            if total_score < min_score:
                min_score = total_score
                best_img = img

        if best_img:
            cap["matched_img_id"] = best_img["id"]
            best_img["matched_caption"] = cap["label"]

    # 3. 텍스트/이미지 추출
    extracted_images_map = {}
    for page_num, page in enumerate(doc):
        page_items = []
        text_blocks = page.get_text("blocks")
        for b in text_blocks:
            bbox = fitz.Rect(b[0], b[1], b[2], b[3])
            matched_cap = next((c for c in all_captions if c["page"] == page_num and c["bbox"] == bbox), None)
            text = b[4]
            if matched_cap and matched_cap["matched_img_id"]:
                text = text.strip() + f"\n[SYSTEM: Matches <<<<{matched_cap['matched_img_id']}>>>>]\n"
            page_items.append({"type": "text", "y0": b[1], "x0": b[0], "text": text})

        page_imgs = [img for img in all_images_info if img["page"] == page_num]
        for img_info in page_imgs:
            rect = img_info["bbox"]
            padding = 35
            clip_rect = fitz.Rect(rect.x0 - padding, rect.y0 - padding, rect.x1 + padding,
                                  rect.y1 + padding) & page.rect
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat, clip=clip_rect)
            img_bytes = pix.tobytes("png")

            img_id = img_info["id"]
            initial_label = img_info["matched_caption"] if img_info["matched_caption"] else "Unknown"

            marker_text = f"\n<<<<{img_id}>>>>\n"
            if img_info["matched_caption"]:
                marker_text = f"\n<<<<{img_id} (Matched with {initial_label})>>>>\n"

            page_items.append({
                "type": "image", "y0": rect.y0, "x0": rect.x0,
                "text": marker_text,
                "id": img_id, "bytes": img_bytes, "page": page_num + 1
            })

            if img_id not in extracted_images_map:
                extracted_images_map[img_id] = {
                    "id": img_id, "page": page_num + 1, "bytes": img_bytes,
                    "initial_label": initial_label
                }

        page_items.sort(key=lambda item: (item["y0"], item["x0"]))
        for item in page_items: final_text_content += item["text"]

    extracted_images = list(extracted_images_map.values())
    return final_text_content, extracted_images


def get_gemini_analysis(text, total_images):
    prompt = f"""
    너는 논문 분석 전문가야. 아래 텍스트를 읽고 JSON으로 추출해.

    [지시사항]
    1. **모든 내용은 한국어로 번역.**
    2. 요약(summary)은 '최소 2문장 ~ 최대 5문장' 사이로 작성.
    3. **이미지 매칭 시, 텍스트에 있는 `(Matched with ...)` 정보를 최우선으로 따를 것.**

    [요청 항목]
    0. title, author, affiliation, year, purpose
    1. 요약 (intro, body, conclusion)
    2. key_images_desc, referenced_images

    [출력 포맷 JSON]
    {{
        "title": "...",
        "author": "...", "affiliation": "...", "year": "...", "purpose": "...",
        "intro_summary": "- ...", "body_summary": "- ...", "conclusion_summary": "- ...",
        "key_images_desc": "...",
        "referenced_images": [ {{ "img_id": "Image_5", "real_label": "Figure 1", "caption": "설명" }} ]
    }}

    [텍스트]:
    """ + text[:50000]

    try:
        response = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
        return json.loads(response.text)
    except Exception as e:
        return {"error": str(e)}


def create_excel(paper_number, analysis_json, images, final_figures, final_tables):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    header_style = workbook.add_format(
        {'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center',
         'valign': 'vcenter'})
    content_style = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
    fig_style = workbook.add_format(
        {'bold': True, 'valign': 'center', 'border': 1, 'bg_color': '#E7E6E6', 'align': 'center'})
    tbl_style = workbook.add_format(
        {'bold': True, 'valign': 'center', 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})

    ws1 = workbook.add_worksheet("논문 핵심 분석")
    ws1.set_column('A:A', 25)
    ws1.set_column('B:B', 90)

    data_map = [
        ("No.", paper_number),
        ("논문 제목", analysis_json.get('title', '제목 없음')),
        ("저자", analysis_json.get('author', '-')),
        ("저자 소속", analysis_json.get('affiliation', '-')),
        ("발행년도", analysis_json.get('year', '-')),
        ("연구 목적", analysis_json.get('purpose', '-')),
        ("서론 요약", analysis_json.get('intro_summary', '-')),
        ("본론 요약", analysis_json.get('body_summary', '-')),
        ("결론 요약", analysis_json.get('conclusion_summary', '-')),
        ("주요 표/그림 설명", analysis_json.get('key_images_desc', '-'))
    ]

    ws1.write(0, 0, "항목", header_style)
    ws1.write(0, 1, "내용", header_style)

    current_row = 1
    for label, content in data_map:
        ws1.write(current_row, 0, label, header_style)
        ws1.write(current_row, 1, content, content_style)
        current_row += 1

    # Figure 섹션
    if final_figures:
        current_row += 1
        ws1.write(current_row, 0, "Figures (그림)", header_style)
        ws1.write(current_row, 1, "▼ 주요 그림 목록", header_style)
        current_row += 1
        if current_row % 2 != 0: current_row += 1
        for item in final_figures:
            _write_row_dynamic(ws1, item, images, current_row, fig_style, content_style)
            current_row += 2

            # Table 섹션
    if final_tables:
        current_row += 1
        ws1.write(current_row, 0, "Tables (표)", header_style)
        ws1.write(current_row, 1, "▼ 주요 표 목록", header_style)
        current_row += 1
        if current_row % 2 != 0: current_row += 1
        for item in final_tables:
            _write_row_dynamic(ws1, item, images, current_row, tbl_style, content_style)
            current_row += 2

    workbook.close()
    output.seek(0)
    return output


def _write_row_dynamic(ws, item, images, row, label_fmt, content_fmt):
    clean_id = normalize_id(item.get('img_id'))
    target = next((img for img in images if img['id'] == clean_id), None)

    ws.write(row, 0, item.get('real_label'), label_fmt)
    ws.write(row, 1, f"📄 {item.get('caption')}", content_fmt)

    img_row = row + 1

    if target:
        try:
            with Image.open(io.BytesIO(target['bytes'])) as img:
                w_px, h_px = img.size

            base_scale = 0.5
            display_h_px = h_px * base_scale
            row_height_pt = display_h_px * 0.75

            MAX_EXCEL_HEIGHT = 400
            final_scale = base_scale

            if row_height_pt > MAX_EXCEL_HEIGHT:
                row_height_pt = MAX_EXCEL_HEIGHT
                final_scale = (MAX_EXCEL_HEIGHT / 0.75) / h_px

            ws.set_row(img_row, row_height_pt)

            ws.insert_image(img_row, 1, f"{clean_id}.png", {
                'image_data': io.BytesIO(target['bytes']),
                'x_scale': final_scale,
                'y_scale': final_scale,
                'x_offset': 0, 'y_offset': 0,
                'object_position': 1
            })
        except:
            pass


# -----------------------------------------------------------
# [6] 실행 로직
# -----------------------------------------------------------

if 'analyzed_data' not in st.session_state:
    st.session_state.analyzed_data = None

paper_num = st.text_input("1. 논문 번호 입력", value="1")
uploaded_file = st.file_uploader("2. PDF 파일 업로드", type="pdf")

if uploaded_file and paper_num:
    if st.session_state.analyzed_data and st.session_state.analyzed_data['file_name'] != uploaded_file.name:
        st.session_state.analyzed_data = None

    if st.button("분석 및 엑셀 변환 시작"):
        if st.session_state.analyzed_data and st.session_state.analyzed_data['file_name'] == uploaded_file.name:
            st.success("⚡ 저장된 분석 결과를 불러옵니다.")
        else:
            with st.spinner(f"[{SELECTED_MODEL_NAME}] 분석 중..."):
                try:
                    text, images = extract_data_from_pdf(uploaded_file)

                    # [수정됨] Vision OCR 과정 제거 및 바로 Gemini 분석 요청
                    result = get_gemini_analysis(text, len(images))

                    if "error" in result:
                        st.error(f"오류: {result['error']}")
                    else:
                        ref_imgs = result.get('referenced_images', [])
                        final_figs, final_tbls = [], []

                        # [분류 로직] Gemini의 텍스트 분석 결과(real_label)에만 의존
                        for item in ref_imgs:
                            label = item.get('real_label', 'Figure')

                            # 'Table' 또는 '표'라는 단어가 들어가면 표로 분류
                            if "Table" in label or "표" in label:
                                final_tbls.append(item)
                            else:
                                final_figs.append(item)


                        def sort_key(x):
                            nums = re.findall(r'\d+', x.get('real_label', '0'))
                            return int(nums[0]) if nums else 999


                        final_figs.sort(key=sort_key)
                        final_tbls.sort(key=sort_key)

                        st.session_state.analyzed_data = {
                            'file_name': uploaded_file.name,
                            'json': result,
                            'images': images,
                            'figs': final_figs,
                            'tbls': final_tbls
                        }
                        st.success("완료! 분석이 끝났습니다.")

                except Exception as e:
                    st.error(f"오류: {e}")

    if st.session_state.analyzed_data:
        data = st.session_state.analyzed_data
        excel_data = create_excel(paper_num, data['json'], data['images'], data['figs'], data['tbls'])

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name=f"Analysis_v5.9_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )