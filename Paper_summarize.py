import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import xlsxwriter
import io
import json
import re
from PIL import Image

# -----------------------------------------------------------
# [1] 페이지 설정
# -----------------------------------------------------------
st.set_page_config(page_title="논문 분석 Pro", layout="wide")

# -----------------------------------------------------------
# [2] 메인 UI
# -----------------------------------------------------------
st.title("📑 논문 분석 Pro [ver7.8 - Width Detection]")
st.caption("✅ 글자 수뿐만 아니라 '텍스트 너비'로 본문/라벨 구분 | 본문 침범 완벽 방지")

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

        preferred = ['gemini-2.5-flash', 'gemini-1.5-flash']
        available_models.sort(key=lambda x: (x not in preferred, x))

        selected_model_name = st.selectbox(
            "✅ 모델 선택 (2.5-flash 기본)",
            available_models,
            index=0
        )
        SELECTED_MODEL_NAME = f"models/{selected_model_name}"
        st.success(f"연결됨: {selected_model_name}")

    except Exception as e:
        st.error(f"모델 목록 오류: {e}")
        st.stop()

model = genai.GenerativeModel(SELECTED_MODEL_NAME)


# -----------------------------------------------------------
# [4] 유틸리티 함수
# -----------------------------------------------------------
def normalize_id(ref_text):
    nums = re.findall(r'\d+', str(ref_text))
    return f"Image_{nums[0]}" if nums else None


def standardize_label_to_korean(label_text):
    if not label_text: return ("Unknown", 999, "미분류")

    label_upper = str(label_text).upper()
    detected_type = "Figure"
    korean_prefix = "그림"

    if "TAB" in label_upper or "표" in label_upper:
        detected_type = "Table"
        korean_prefix = "표"
    elif "FIG" in label_upper or "그림" in label_upper:
        detected_type = "Figure"
        korean_prefix = "그림"

    nums = re.findall(r'\d+', label_text)
    if nums:
        detected_num = int(nums[0])
        final_label = f"{korean_prefix} {detected_num}"
    else:
        detected_num = 999
        final_label = f"{korean_prefix} (번호 없음)"

    return (detected_type, detected_num, final_label)


# -----------------------------------------------------------
# [5] 핵심 로직 함수 (너비 기반 장벽 감지 추가)
# -----------------------------------------------------------
def extract_data_from_pdf(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    final_text_content = ""
    image_counter = 1

    all_page_images = []
    extracted_images_map = {}

    for page_num, page in enumerate(doc):
        blocks = page.get_text("blocks")
        blocks.sort(key=lambda b: b[1])
        final_text_content += page.get_text() + "\n"

        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_data = Image.open(io.BytesIO(pix.tobytes("png")))
        all_page_images.append(img_data)

        # 시각적 요소(Visual Elements) 수집
        visual_elements = []
        drawings = page.get_drawings()
        for d in drawings:
            visual_elements.append(d["rect"])

        image_list = page.get_images(full=True)
        for img in image_list:
            xref = img[0]
            try:
                rects = page.get_image_rects(xref)
                for r in rects:
                    visual_elements.append(r)
            except:
                pass

        page_width = page.rect.width
        page_height = page.rect.height
        page_center_x = page_width / 2

        for i, block in enumerate(blocks):
            text = block[4].strip()
            bbox = fitz.Rect(block[0], block[1], block[2], block[3])

            if len(text) < 300 and re.search(r"(Fig|Figure|Table|그림|표)\s*[\.|\s]\s*\d+", text, re.IGNORECASE):

                is_table = "Table" in text or "표" in text or "TABLE" in text.upper()
                label_match = re.search(r"(Fig\.?|Figure|Table|그림|표)\s*\d+", text, re.IGNORECASE)
                real_label = label_match.group(0) if label_match else text[:15]

                # 단(Column) 판단 및 너비 계산
                caption_center_x = (bbox.x0 + bbox.x1) / 2

                if (bbox.x1 - bbox.x0) > (page_width * 0.6):  # 1단 (Full Width)
                    col_x0, col_x1 = 0, page_width
                    col_width = page_width
                elif caption_center_x < page_center_x:  # 왼쪽 단
                    col_x0, col_x1 = 0, page_center_x + 20
                    col_width = page_center_x
                else:  # 오른쪽 단
                    col_x0, col_x1 = page_center_x - 20, page_width
                    col_width = page_width - page_center_x

                crop_rect = None

                # --- [A] Table (캡션 아래) ---
                if is_table:
                    top_y = max(0, bbox.y0 - 5)
                    search_limit_y = min(page_height, top_y + 600)

                    barrier_y = search_limit_y

                    for other_block in blocks:
                        if other_block == block: continue
                        o_bbox = fitz.Rect(other_block[0], other_block[1], other_block[2], other_block[3])

                        if (o_bbox.x1 > col_x0 and o_bbox.x0 < col_x1) and (o_bbox.y0 > bbox.y1 + 5):
                            other_text = other_block[4].strip()
                            o_width = o_bbox.x1 - o_bbox.x0

                            # [핵심 수정] 글자 수가 많거나(50자), 너비가 단의 80% 이상이면 장벽
                            is_wide_block = o_width > (col_width * 0.8)
                            is_long_text = len(other_text) > 50

                            if is_long_text or is_wide_block:
                                if o_bbox.y0 < barrier_y:
                                    barrier_y = o_bbox.y0
                                    break

                                    # Visual Element 보정
                    max_visual_y = top_y + 50
                    for v_rect in visual_elements:
                        if (v_rect.x1 > col_x0 and v_rect.x0 < col_x1) and \
                                (v_rect.y0 >= bbox.y1) and (v_rect.y1 <= barrier_y):
                            if v_rect.y1 > max_visual_y:
                                max_visual_y = v_rect.y1

                    if max_visual_y == top_y + 50 and barrier_y == search_limit_y:
                        final_bottom = min(page_height, top_y + 400)
                    else:
                        final_bottom = max(max_visual_y, barrier_y if barrier_y < search_limit_y else top_y + 100)
                        if barrier_y < search_limit_y:
                            final_bottom = min(final_bottom, barrier_y)

                    crop_rect = fitz.Rect(col_x0, top_y, col_x1, final_bottom)

                # --- [B] Figure (캡션 위) ---
                else:
                    bottom_y = min(page_height, bbox.y1 + 5)
                    search_limit_y = max(0, bottom_y - 600)

                    barrier_y = search_limit_y

                    for other_block in blocks:
                        if other_block == block: continue
                        o_bbox = fitz.Rect(other_block[0], other_block[1], other_block[2], other_block[3])

                        if (o_bbox.x1 > col_x0 and o_bbox.x0 < col_x1) and (o_bbox.y1 < bbox.y0 - 5):
                            other_text = other_block[4].strip()
                            o_width = o_bbox.x1 - o_bbox.x0

                            # [핵심 수정] 너비가 넓으면 무조건 본문으로 간주 (Figure 2 문제 해결)
                            is_wide_block = o_width > (col_width * 0.8)
                            is_long_text = len(other_text) > 50

                            if is_long_text or is_wide_block:
                                if o_bbox.y1 > barrier_y:
                                    barrier_y = o_bbox.y1

                    # Visual Element 보정
                    min_visual_y = bottom_y - 50
                    found_visual = False
                    for v_rect in visual_elements:
                        if (v_rect.x1 > col_x0 and v_rect.x0 < col_x1) and \
                                (v_rect.y1 <= bbox.y0) and (v_rect.y0 >= barrier_y):
                            if v_rect.y0 < min_visual_y:
                                min_visual_y = v_rect.y0
                                found_visual = True

                    # 시각적 요소가 발견되면 그 위주로, 없으면 장벽까지
                    if found_visual:
                        # 시각적 요소 위쪽을 쓰되, 장벽보단 아래여야 함
                        final_top = max(min_visual_y, barrier_y)
                    else:
                        # 시각적 요소 못 찾음 -> 장벽이 있으면 장벽까지, 없으면 400px
                        if barrier_y > search_limit_y:
                            final_top = barrier_y
                        else:
                            final_top = max(0, bottom_y - 400)

                    crop_rect = fitz.Rect(col_x0, final_top, col_x1, bottom_y)

                # --- 3. 캡처 실행 ---
                if crop_rect:
                    if crop_rect.height < 50:
                        if is_table:
                            crop_rect.y1 += 200
                        else:
                            crop_rect.y0 -= 200

                    try:
                        clip_pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=crop_rect)

                        # [필터] 아주 작은 노이즈만 제거
                        if clip_pix.width < 50 or clip_pix.height < 50:
                            continue

                        img_bytes = clip_pix.tobytes("png")
                        img_id = f"Image_{image_counter}"
                        image_counter += 1

                        extracted_images_map[img_id] = {
                            "id": img_id,
                            "page": page_num + 1,
                            "bytes": img_bytes,
                            "initial_label": text,
                            "real_label": real_label
                        }
                    except Exception as e:
                        print(f"Crop Error: {e}")
                        continue

    extracted_images = list(extracted_images_map.values())
    return final_text_content, extracted_images, all_page_images


def get_gemini_analysis(text, total_images, all_page_images):
    inputs = []

    prompt = f"""
    너는 한국어 논문 분석 전문가야. 제공된 자료를 보고 JSON을 추출해.

    [절대 규칙]
    1. **모든 요약은 반드시 '한국어(Korean)'로 작성.**
    2. **요약은 '개조식(Bullet Points)'으로 작성.**
    3. **이미지 매칭:** `referenced_images`의 `real_label`은 텍스트 번호(예: 그림 1)와 일치해야 함.

    [요청 항목]
    0. title, author, affiliation, year, purpose
    1. 요약 (intro_summary, body_summary, conclusion_summary)
    2. key_images_desc
    3. referenced_images (img_id, real_label, caption)

    [출력 포맷 JSON]
    {{
        "title": "...",
        "author": "...", "affiliation": "...", "year": "...", "purpose": "...",
        "intro_summary": "- ...", 
        "body_summary": "- ...", 
        "conclusion_summary": "- ...",
        "key_images_desc": "...",
        "referenced_images": [ 
            {{ "img_id": "Image_1", "real_label": "Figure 1", "caption": "설명" }}
        ]
    }}
    """

    inputs.append(prompt)

    is_text_valid = len(text.strip()) > 500

    if is_text_valid:
        inputs.append(f"[추출된 텍스트 데이터]:\n{text[:50000]}")
    else:
        inputs.append("[시스템 알림: 텍스트 추출 실패. 아래의 '전체 페이지 이미지'를 읽고 분석하세요.]")

    if not is_text_valid:
        max_pages = 30
        for i, img in enumerate(all_page_images[:max_pages]):
            inputs.append(f"Page {i + 1} Image:")
            inputs.append(img)

    try:
        response = model.generate_content(inputs, generation_config={"response_mime_type": "application/json"})
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
        ("논문 제목", analysis_json.get('title', '-')),
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
        if isinstance(content, list):
            content = "\n".join(map(str, content))
        elif content is None:
            content = "-"
        ws1.write(current_row, 0, label, header_style)
        ws1.write(current_row, 1, str(content), content_style)
        current_row += 1

    if final_figures:
        current_row += 1
        ws1.write(current_row, 0, "그림 (Figures)", header_style)
        ws1.write(current_row, 1, "▼ 주요 그림 목록", header_style)
        current_row += 1
        if current_row % 2 != 0: current_row += 1
        for item in final_figures:
            _write_row_dynamic(ws1, item, images, current_row, fig_style, content_style)
            current_row += 2

    if final_tables:
        current_row += 1
        ws1.write(current_row, 0, "표 (Tables)", header_style)
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

    final_label = item.get('korean_label', item.get('real_label', '그림'))
    caption_text = item.get('caption', '설명 없음')

    ws.write(row, 0, str(final_label), label_fmt)
    ws.write(row, 1, f"📄 {str(caption_text)}", content_fmt)

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
            with st.spinner(f"[{SELECTED_MODEL_NAME}] 분석 중... (너비 기반 본문 감지)"):
                try:
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if len(text.strip()) < 500:
                        st.warning("⚠️ 텍스트 추출이 불안정하여 전체 페이지 분석을 병행합니다.")
                    else:
                        st.info(f"✅ 텍스트 및 {len(images)}개의 이미지(본문 제외) 추출 완료!")

                    result = get_gemini_analysis(text, len(images), all_page_imgs)

                    if "error" in result:
                        st.error(f"AI 분석 오류: {result['error']}")
                    else:
                        ref_imgs = result.get('referenced_images', [])
                        final_figs, final_tbls = [], []

                        for item in ref_imgs:
                            raw_label = item.get('real_label', 'Unknown')
                            detected_type, detected_num, korean_label = standardize_label_to_korean(raw_label)

                            item['sort_num'] = detected_num
                            item['korean_label'] = korean_label

                            if detected_type == "Table":
                                final_tbls.append(item)
                            else:
                                final_figs.append(item)

                        final_figs.sort(key=lambda x: x['sort_num'])
                        final_tbls.sort(key=lambda x: x['sort_num'])

                        st.session_state.analyzed_data = {
                            'file_name': uploaded_file.name,
                            'json': result,
                            'images': images,
                            'figs': final_figs,
                            'tbls': final_tbls
                        }
                        st.success("완료! 분석이 끝났습니다.")

                except Exception as e:
                    st.error(f"시스템 오류: {e}")

    if st.session_state.analyzed_data:
        data = st.session_state.analyzed_data
        excel_data = create_excel(paper_num, data['json'], data['images'], data['figs'], data['tbls'])

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name=f"Analysis_v7.8_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
