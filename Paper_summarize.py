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
st.title("📑 논문 분석 Pro [ver6.6 - Smart Crop]")
st.caption("✅ 로고/아이콘 자동 제거 | 캡션 위치 기반 '영역 캡처'로 정확도 향상 | 요약 한글 필수")

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
    """
    라벨을 분석해서 한글로 변환 (Figure 1 -> 그림 1)
    """
    if not label_text:
        return ("Unknown", 999, "미분류")

    label_upper = str(label_text).upper()

    # 1. 타입 결정 및 한글 변환
    detected_type = "Figure"
    korean_prefix = "그림"

    if "TAB" in label_upper or "표" in label_upper:
        detected_type = "Table"
        korean_prefix = "표"
    elif "FIG" in label_upper or "그림" in label_upper:
        detected_type = "Figure"
        korean_prefix = "그림"

    # 2. 번호 추출
    nums = re.findall(r'\d+', label_text)
    if nums:
        detected_num = int(nums[0])
        final_label = f"{korean_prefix} {detected_num}"
    else:
        detected_num = 999
        final_label = f"{korean_prefix} (번호 없음)"

    return (detected_type, detected_num, final_label)


# -----------------------------------------------------------
# [5] 핵심 로직 함수 (완전히 새로 작성됨)
# -----------------------------------------------------------
def extract_data_from_pdf(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    final_text_content = ""
    image_counter = 1

    all_page_images = []  # 텍스트 인식 실패 대비용
    extracted_images_map = {}  # 최종 이미지 저장소

    for page_num, page in enumerate(doc):
        # 1. 텍스트 추출
        text_blocks = page.get_text("blocks")
        for b in text_blocks:
            final_text_content += b[4].strip() + "\n"

        # 2. 페이지 전체 이미지 저장 (AI 분석용)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_data = Image.open(io.BytesIO(pix.tobytes("png")))
        all_page_images.append(img_data)

        # 3. [핵심] 캡션 텍스트 찾기 (Fig, Table)
        # 텍스트 블록을 순회하며 "Fig"나 "Table"로 시작하는 줄을 찾음
        captions = []
        for b in text_blocks:
            text = b[4].strip()
            # 캡션 조건: Fig/Table로 시작하고 길이가 너무 길지 않은 것
            if re.match(r"^(Fig|Figure|Table|그림|표)\s*\.?\s*\d+", text, re.IGNORECASE) and len(text) < 300:
                bbox = fitz.Rect(b[0], b[1], b[2], b[3])
                captions.append({"text": text, "bbox": bbox})

        # 4. [핵심] 캡션 기준으로 영역 캡처 (Smart Crop)
        # 이미지를 찾는 게 아니라, 캡션 위치를 기준으로 화면을 잘라버림
        for cap in captions:
            text = cap["text"]
            bbox = cap["bbox"]

            is_table = "Table" in text or "표" in text
            img_id = f"Image_{image_counter}"
            image_counter += 1

            # 잘라낼 영역 계산 (Crop Area)
            page_rect = page.rect
            crop_rect = None

            if is_table:
                # 표는 캡션이 보통 '위'에 있음 -> 캡션 '아래'를 캡처
                # 캡션 y1부터 페이지 끝 혹은 적당한 높이(300~400px)까지
                crop_rect = fitz.Rect(page_rect.x0 + 30, bbox.y1, page_rect.x1 - 30,
                                      min(bbox.y1 + 400, page_rect.y1 - 50))
            else:
                # 그림은 캡션이 보통 '아래'에 있음 -> 캡션 '위'를 캡처
                # 캡션 y0에서 위로 300~400px 정도
                crop_rect = fitz.Rect(page_rect.x0 + 30, max(bbox.y0 - 400, page_rect.y0 + 50), page_rect.x1 - 30,
                                      bbox.y0)

            # 영역 캡처 실행
            try:
                clip_pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=crop_rect)

                # [중요] 캡처한 이미지가 너무 단색이거나(흰색) 작으면 버림 (빈 공간 방지)
                if clip_pix.width < 50 or clip_pix.height < 50:
                    continue

                img_bytes = clip_pix.tobytes("png")

                # 라벨 정규화 (Fig. 1 -> 그림 1)
                label_match = re.match(r"(Fig\.?|Figure|Table|그림|표)\s*\d+", text, re.IGNORECASE)
                real_label = label_match.group(0) if label_match else text[:10]

                extracted_images_map[img_id] = {
                    "id": img_id,
                    "page": page_num + 1,
                    "bytes": img_bytes,
                    "initial_label": text,  # 전체 캡션
                    "real_label": real_label  # 그림 1
                }
            except Exception as e:
                print(f"Crop Error: {e}")
                continue

    extracted_images = list(extracted_images_map.values())
    return final_text_content, extracted_images, all_page_images


def get_gemini_analysis(text, total_images, all_page_images):
    inputs = []

    # [프롬프트 강화] 한국어 요약 필수 & 이미지 매칭 지시
    prompt = f"""
    너는 한국어 논문 분석 전문가야. 제공된 자료를 보고 JSON을 추출해.

    [절대 규칙]
    1. **모든 요약(Summary)은 반드시 '한국어(Korean)'로 번역해서 작성해.** (영어 내용 금지)
    2. **요약은 '개조식(Bullet Points)'으로 간결하게 작성해.**
    3. **이미지 매칭:**
       - 내가 추출한 이미지 리스트(`referenced_images`)에 있는 `real_label` (예: 그림 1)과 내용을 매칭해서 설명해.
       - 엉뚱한 이미지를 매칭하지 마.

    [요청 항목]
    0. title, author, affiliation, year, purpose
    1. 요약 (intro_summary, body_summary, conclusion_summary) - **한국어 필수**
    2. key_images_desc - **한국어 필수**
    3. referenced_images (이미지 ID와 설명을 유지해)

    [출력 포맷 JSON]
    {{
        "title": "...",
        "author": "...", "affiliation": "...", "year": "...", "purpose": "...",
        "intro_summary": "- ...", 
        "body_summary": "- ...", 
        "conclusion_summary": "- ...",
        "key_images_desc": "...",
        "referenced_images": [ 
            {{ "img_id": "Image_1", "real_label": "Figure 1", "caption": "한국어 설명" }}
        ]
    }}
    """

    inputs.append(prompt)

    is_text_valid = len(text.strip()) > 500

    if is_text_valid:
        inputs.append(f"[추출된 텍스트 데이터]:\n{text[:50000]}")
    else:
        inputs.append("[시스템 알림: 텍스트 추출 실패. 아래의 '전체 페이지 이미지'를 보고 내용을 요약하세요.]")

    # 텍스트가 부족할 때만 전체 페이지 이미지 전송 (비용/속도 절약)
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

    # Figure 섹션
    if final_figures:
        current_row += 1
        ws1.write(current_row, 0, "그림 (Figures)", header_style)
        ws1.write(current_row, 1, "▼ 주요 그림 목록", header_style)
        current_row += 1
        if current_row % 2 != 0: current_row += 1
        for item in final_figures:
            _write_row_dynamic(ws1, item, images, current_row, fig_style, content_style)
            current_row += 2

    # Table 섹션
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

    # 한글 라벨 적용
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
            with st.spinner(f"[{SELECTED_MODEL_NAME}] 분석 중... (영역 캡처 모드)"):
                try:
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if len(text.strip()) < 500:
                        st.warning("⚠️ 텍스트 추출이 불안정하여 전체 페이지 분석을 병행합니다.")
                    else:
                        st.info(f"✅ 텍스트 및 {len(images)}개의 주요 영역(Fig/Table) 추출 완료!")

                    result = get_gemini_analysis(text, len(images), all_page_imgs)

                    if "error" in result:
                        st.error(f"AI 분석 오류: {result['error']}")
                    else:
                        ref_imgs = result.get('referenced_images', [])
                        final_figs, final_tbls = [], []

                        for item in ref_imgs:
                            raw_label = item.get('real_label', 'Unknown')

                            # 한글 변환
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
                        st.success("완료! 로고는 버리고, 진짜 그림과 표만 가져왔습니다.")

                except Exception as e:
                    st.error(f"시스템 오류: {e}")

    if st.session_state.analyzed_data:
        data = st.session_state.analyzed_data
        excel_data = create_excel(paper_num, data['json'], data['images'], data['figs'], data['tbls'])

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name=f"Analysis_v6.6_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
