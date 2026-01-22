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
st.title("📑 논문 분석 Pro [ver7.1 - 2-Column Expert]")
st.caption("✅ 2단 레이아웃 완벽 대응 | 중앙선 기준 좌우 분리 캡처 | 텍스트 뭉치 회피")

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
# [5] 핵심 로직 함수 (2단 레이아웃 대응)
# -----------------------------------------------------------
def extract_data_from_pdf(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    final_text_content = ""
    image_counter = 1

    all_page_images = []
    extracted_images_map = {}

    for page_num, page in enumerate(doc):
        # 1. 텍스트 블록 정보 수집
        # blocks: (x0, y0, x1, y1, "text", block_no, block_type)
        blocks = page.get_text("blocks")

        # 전체 텍스트 추출 (요약용)
        final_text_content += page.get_text() + "\n"

        # AI 분석용 페이지 이미지
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_data = Image.open(io.BytesIO(pix.tobytes("png")))
        all_page_images.append(img_data)

        # 페이지 중앙 좌표 계산 (양단 판단용)
        page_width = page.rect.width
        page_center_x = page_width / 2

        # 2. 캡션 탐색 및 영역 계산
        for i, block in enumerate(blocks):
            text = block[4].strip()
            bbox = fitz.Rect(block[0], block[1], block[2], block[3])  # 캡션의 위치

            # 캡션 조건: Fig/Table로 시작하고 짧은 문장
            if len(text) < 300 and re.match(r"^(Fig|Figure|Table|그림|표)\s*[\.|\s]\s*\d+", text, re.IGNORECASE):

                is_table = "Table" in text or "표" in text or "TABLE" in text.upper()

                # 라벨 추출
                label_match = re.match(r"(Fig\.?|Figure|Table|그림|표)\s*\d+", text, re.IGNORECASE)
                real_label = label_match.group(0) if label_match else text[:15]

                # --- [중요] 2단 레이아웃 판단 로직 ---
                # 캡션의 중심이 왼쪽에 있는지 오른쪽에 있는지 확인
                caption_center_x = (bbox.x0 + bbox.x1) / 2

                # 검색할 X축 범위 설정 (좌/우 분리)
                if caption_center_x < page_center_x:
                    # 왼쪽 단
                    search_x_min = 0
                    search_x_max = page_center_x + 10  # 약간의 여유
                else:
                    # 오른쪽 단
                    search_x_min = page_center_x - 10
                    search_x_max = page_width

                # 만약 캡션 자체가 페이지 너비의 70% 이상을 차지하면 '1단(통짜)'로 간주
                if (bbox.x1 - bbox.x0) > (page_width * 0.7):
                    search_x_min = 0
                    search_x_max = page_width

                crop_rect = None

                # --- [A] Table 로직 (캡션 아래 검색) ---
                if is_table:
                    top_y = bbox.y1  # 캡션 바로 아래
                    bottom_y = page.rect.y1 - 30  # 기본값: 페이지 끝

                    # 같은 단(Column)에 있는 '다음 텍스트 블록' 찾기 (장벽 감지)
                    closest_next_block_y = bottom_y

                    for other_block in blocks:
                        # 자기 자신은 제외
                        if other_block == block: continue

                        o_bbox = fitz.Rect(other_block[0], other_block[1], other_block[2], other_block[3])

                        # 1. 같은 단(X축 범위)에 있어야 함
                        if not (o_bbox.x1 > search_x_min and o_bbox.x0 < search_x_max):
                            continue

                        # 2. 캡션보다 아래에 있어야 함 (Table 내용보다 더 아래)
                        if o_bbox.y0 > top_y + 10:  # +10은 여유
                            if o_bbox.y0 < closest_next_block_y:
                                closest_next_block_y = o_bbox.y0  # 가장 가까운 아래 블록 갱신

                    bottom_y = closest_next_block_y
                    crop_rect = fitz.Rect(search_x_min, top_y, search_x_max, bottom_y)

                # --- [B] Figure 로직 (캡션 위 검색) ---
                else:
                    bottom_y = bbox.y0  # 캡션 바로 위
                    top_y = page.rect.y0 + 30  # 기본값: 페이지 시작

                    # 같은 단(Column)에 있는 '이전 텍스트 블록' 찾기 (장벽 감지)
                    closest_prev_block_y = top_y

                    for other_block in blocks:
                        if other_block == block: continue

                        o_bbox = fitz.Rect(other_block[0], other_block[1], other_block[2], other_block[3])

                        # 1. 같은 단(X축 범위)에 있어야 함
                        if not (o_bbox.x1 > search_x_min and o_bbox.x0 < search_x_max):
                            continue

                        # 2. 캡션보다 위에 있어야 함
                        if o_bbox.y1 < bottom_y - 5:
                            if o_bbox.y1 > closest_prev_block_y:
                                closest_prev_block_y = o_bbox.y1  # 가장 가까운 위 블록 갱신

                    top_y = closest_prev_block_y
                    crop_rect = fitz.Rect(search_x_min, top_y, search_x_max, bottom_y)

                # --- 3. 이미지 캡처 및 저장 ---
                if crop_rect and crop_rect.height > 30:
                    try:
                        clip_pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=crop_rect)

                        # [필터] 30px 이하는 버림
                        if clip_pix.width < 30 or clip_pix.height < 30:
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
    1. **모든 요약(Summary)은 반드시 '한국어(Korean)'로 작성해.**
    2. **요약은 '개조식(Bullet Points)'으로 작성해.**
    3. **이미지 매칭:**
       - `referenced_images`의 `real_label`은 텍스트에 있는 번호(예: 그림 1, Table 2)와 정확히 일치해야 해.
       - 내가 잘라낸 이미지(`Image_X`)가 해당 그림 번호의 내용과 맞는지 확인해.

    [요청 항목]
    0. title, author, affiliation, year, purpose
    1. 요약 (intro_summary, body_summary, conclusion_summary) - **한국어 필수**
    2. key_images_desc - **한국어 필수**
    3. referenced_images (이미지 ID와 한글 라벨)

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

    ws.write(row, 0, final_label, label_fmt)
    ws.write(row, 1, f"📄 {caption_text}", content_fmt)

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
            with st.spinner(f"[{SELECTED_MODEL_NAME}] 분석 중... (2단 레이아웃 처리)"):
                try:
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if len(text.strip()) < 500:
                        st.warning("⚠️ 텍스트 추출이 불안정하여 전체 페이지 분석을 병행합니다.")
                    else:
                        st.info(f"✅ 텍스트 및 {len(images)}개의 이미지(단 분리 적용) 추출 완료!")

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
            file_name=f"Analysis_v7.1_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
