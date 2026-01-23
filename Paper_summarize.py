import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import xlsxwriter
import io
import json
from PIL import Image

# -----------------------------------------------------------
# [1] 페이지 설정
# -----------------------------------------------------------
st.set_page_config(page_title="논문 분석 Pro", layout="wide")

# -----------------------------------------------------------
# [2] 메인 UI
# -----------------------------------------------------------
st.title("📑 논문 분석 Pro [ver10.1 - Vision + Custom Model]")
st.caption("✅ 딥러닝 비전 인식(좌표 추출) | 모델 선택 기능 복구 (2.5-flash 등 자유 선택)")

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

        # 사용자가 선호했던 순서대로 정렬 (2.5-flash 우선)
        preferred = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']
        available_models.sort(key=lambda x: (x not in preferred, x))

        selected_model_name = st.selectbox(
            "✅ 모델 선택 (2.5-flash 기본)",
            available_models,
            index=0
        )
        SELECTED_MODEL_NAME = f"models/{selected_model_name}"
        st.success(f"연결됨: {selected_model_name}")

        # 모델별 팁 표시
        if "pro" in selected_model_name:
            st.info("💡 Pro 모델: 속도는 느리지만 그림 위치를 더 정확하게 찾습니다.")
        else:
            st.info("⚡ Flash 모델: 속도가 빠릅니다.")

    except Exception as e:
        st.error(f"모델 목록 오류: {e}")
        st.stop()

model = genai.GenerativeModel(SELECTED_MODEL_NAME)


# -----------------------------------------------------------
# [4] 핵심 로직: AI Vision을 이용한 좌표 추출
# -----------------------------------------------------------
def detect_regions_with_gemini(page_image):
    """
    페이지 이미지를 Gemini에게 보내서 Figure와 Table의 좌표를 받아옴.
    """
    prompt = """
    Look at this research paper page. 
    Detect all **Figures (charts, diagrams, photos)** and **Tables**.

    [Rules]
    1. Return Bounding Boxes in **normalized coordinates (0 to 1000)**: [ymin, xmin, ymax, xmax].
    2. **Include Captions:** The bounding box MUST include the Figure/Table label (e.g., "Fig. 1", "Table 1") and its description text.
    3. **Group Together:** If a figure has multiple parts (a, b, c) and one caption, group them into ONE bounding box.
    4. **Output Format:** JSON list of objects.

    Example Output:
    [
      {"type": "Figure", "label": "Fig. 1", "box_2d": [100, 50, 400, 500]},
      {"type": "Table", "label": "Table 1", "box_2d": [500, 50, 700, 950]}
    ]
    """

    try:
        response = model.generate_content(
            [prompt, page_image],
            generation_config={"response_mime_type": "application/json"}
        )
        return json.loads(response.text)
    except:
        return []


def extract_data_from_pdf(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    final_text_content = ""
    image_counter = 1

    all_page_images = []
    extracted_images_map = {}

    # 진행률 표시 바
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_pages = len(doc)

    for page_num, page in enumerate(doc):
        # 진행 상황 업데이트
        status_text.text(f"🔍 AI가 {page_num + 1}/{total_pages} 페이지를 보고 있습니다...")
        progress_bar.progress((page_num + 1) / total_pages)

        # 1. 텍스트 추출 (요약용)
        final_text_content += page.get_text() + "\n"

        # 2. 페이지를 이미지로 변환 (AI 분석용)
        # 해상도를 높여야(dpi=200 이상) 작은 글씨도 잘 보임
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data_bytes = pix.tobytes("png")
        pil_image = Image.open(io.BytesIO(img_data_bytes))
        all_page_images.append(pil_image)

        # 3. [Deep Learning] AI에게 좌표 요청
        # 비전 기능이 있는 모델인지 확인 후 요청
        detected_objects = detect_regions_with_gemini(pil_image)

        page_width = page.rect.width
        page_height = page.rect.height

        # 4. AI가 알려준 좌표대로 자르기
        if detected_objects:
            for obj in detected_objects:
                label = obj.get("label", "Unknown")
                box = obj.get("box_2d")  # [ymin, xmin, ymax, xmax] (0~1000)

                if not box: continue

                # 좌표 정규화 (0~1000 -> 실제 PDF 좌표)
                # Gemini Vision은 [ymin, xmin, ymax, xmax] 순서로 줌
                ymin, xmin, ymax, xmax = box

                real_x0 = (xmin / 1000) * page_width
                real_y0 = (ymin / 1000) * page_height
                real_x1 = (xmax / 1000) * page_width
                real_y1 = (ymax / 1000) * page_height

                # 좌표 유효성 검사 및 여유 공간(Padding) 추가
                pad = 10
                crop_rect = fitz.Rect(
                    max(0, real_x0 - pad),
                    max(0, real_y0 - pad),
                    min(page_width, real_x1 + pad),
                    min(page_height, real_y1 + pad)
                )

                if crop_rect.width < 50 or crop_rect.height < 50: continue

                try:
                    # 고해상도 캡처
                    clip_pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=crop_rect)
                    img_bytes = clip_pix.tobytes("png")

                    img_id = f"Image_{image_counter}"
                    image_counter += 1

                    extracted_images_map[img_id] = {
                        "id": img_id,
                        "page": page_num + 1,
                        "bytes": img_bytes,
                        "initial_label": label,  # AI가 읽은 라벨 (예: Fig. 1)
                        "real_label": label
                    }
                except Exception as e:
                    print(f"Crop Error: {e}")
                    continue

    status_text.text("✅ 분석 완료! 엑셀을 생성합니다.")
    progress_bar.empty()

    extracted_images = list(extracted_images_map.values())
    return final_text_content, extracted_images, all_page_images


def get_gemini_analysis(text, total_images, all_page_images):
    prompt = f"""
    너는 논문 분석 전문가야. 아래 텍스트 데이터를 바탕으로 내용을 한국어로 요약해.

    [지시사항]
    1. 요약(intro, body, conclusion)은 반드시 '한국어(Korean)'로 개조식 작성.
    2. `referenced_images`의 `real_label`은 텍스트의 번호(예: Fig 1, Table 1)와 일치시킬 것.
    3. 이미지가 본문 내용에서 어떤 의미를 갖는지 `caption`에 상세히 적어줘.

    [JSON 형식]
    {{
        "title": "제목", "author": "저자", "affiliation": "소속", "year": "연도", "purpose": "목적",
        "intro_summary": "- ...",
        "body_summary": "- ...",
        "conclusion_summary": "- ...",
        "key_images_desc": "주요 그림 설명 요약",
        "referenced_images": [ {{ "img_id": "Image_1", "real_label": "Fig. 1", "caption": "한국어 설명" }} ]
    }}
    """
    inputs = [prompt]
    # 텍스트가 너무 길면 잘라서 보냄
    if len(text.strip()) > 500:
        inputs.append(f"[Text Data]:\n{text[:50000]}")
    else:
        inputs.append("텍스트가 부족합니다. 이미지를 참고하세요.")

    try:
        response = model.generate_content(inputs, generation_config={"response_mime_type": "application/json"})
        return json.loads(response.text)
    except Exception as e:
        return {"error": str(e)}


# -----------------------------------------------------------
# [6] 엑셀 생성 및 유틸리티 (기존과 동일하지만 안정성 강화)
# -----------------------------------------------------------
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

    import re
    nums = re.findall(r'\d+', label_text)
    if nums:
        detected_num = int(nums[0])
        final_label = f"{korean_prefix} {detected_num}"
    else:
        detected_num = 999
        final_label = f"{korean_prefix} (번호 없음)"
    return (detected_type, detected_num, final_label)


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

    def write_section(title, items, style):
        nonlocal current_row
        if not items: return
        current_row += 1
        ws1.write(current_row, 0, title, header_style)
        ws1.write(current_row, 1, f"▼ 주요 {title} 목록", header_style)
        current_row += 1

        for item in items:
            clean_id = item.get('img_id')
            target = next((img for img in images if img['id'] == clean_id), None)

            final_label = item.get('korean_label', item.get('real_label', '그림'))
            caption_text = item.get('caption', '설명 없음')

            ws1.write(current_row, 0, str(final_label), style)
            ws1.write(current_row, 1, f"📄 {str(caption_text)}", content_style)

            img_row = current_row + 1
            if target:
                try:
                    with Image.open(io.BytesIO(target['bytes'])) as img:
                        w_px, h_px = img.size

                    # 이미지 크기 최적화 (엑셀 셀 높이 조절)
                    scale = 0.5
                    display_h = h_px * scale
                    row_h = display_h * 0.75

                    if row_h > 400:
                        row_h = 400
                        scale = (400 / 0.75) / h_px

                    ws1.set_row(img_row, row_h)
                    ws1.insert_image(img_row, 1, f"{clean_id}.png", {
                        'image_data': io.BytesIO(target['bytes']),
                        'x_scale': scale, 'y_scale': scale,
                        'x_offset': 5, 'y_offset': 5, 'object_position': 1
                    })
                except:
                    pass
            current_row += 2

    write_section("그림 (Figures)", final_figures, fig_style)
    write_section("표 (Tables)", final_tables, tbl_style)

    workbook.close()
    output.seek(0)
    return output


# -----------------------------------------------------------
# [7] 실행 로직
# -----------------------------------------------------------

if 'analyzed_data' not in st.session_state:
    st.session_state.analyzed_data = None

paper_num = st.text_input("1. 논문 번호 입력", value="1")
uploaded_file = st.file_uploader("2. PDF 파일 업로드", type="pdf")

if uploaded_file and paper_num:
    if st.session_state.analyzed_data and st.session_state.analyzed_data['file_name'] != uploaded_file.name:
        st.session_state.analyzed_data = None

    if st.button("분석 및 엑셀 변환 시작"):
        # 진행 중 상태 표시
        if st.session_state.analyzed_data and st.session_state.analyzed_data['file_name'] == uploaded_file.name:
            st.success("⚡ 저장된 분석 결과를 불러옵니다.")
        else:
            with st.spinner(f"[{SELECTED_MODEL_NAME}] AI가 눈으로 보고 분석 중... (시간이 조금 걸립니다)"):
                try:
                    # 1. 이미지 추출 (AI Vision 사용)
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if not images:
                        st.warning("⚠️ AI가 그림/표를 찾지 못했습니다. 모델을 '1.5-pro'로 변경해보세요.")
                    else:
                        st.info(f"✅ AI가 {len(images)}개의 그림/표 영역을 인식했습니다!")

                    # 2. 내용 분석
                    result = get_gemini_analysis(text, len(images), all_page_imgs)

                    if "error" in result:
                        st.error(f"AI 분석 오류: {result['error']}")
                    else:
                        # 3. 매칭 및 정렬
                        ref_imgs = result.get('referenced_images', [])

                        final_figs, final_tbls = [], []

                        for img in images:
                            img_label = img['initial_label']  # Vision이 읽은 라벨 (예: Fig 1)

                            # 분석 결과에서 설명 찾기
                            matched_caption = "설명 없음"
                            for ref in ref_imgs:
                                # 단순 포함 관계 확인 (Fig 1 in Figure 1)
                                # AI가 읽은 라벨과 분석된 라벨을 최대한 매칭
                                if normalize_id(img_label) == normalize_id(ref.get('real_label', '')):
                                    matched_caption = ref.get('caption', '-')
                                    break

                            # 분류 및 저장
                            d_type, d_num, k_label = standardize_label_to_korean(img_label)

                            item = {
                                'img_id': img['id'],
                                'real_label': img_label,
                                'korean_label': k_label,
                                'caption': matched_caption,
                                'sort_num': d_num
                            }

                            if d_type == 'Table':
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
                        st.success("완료! AI가 보고 판단한 결과입니다.")

                except Exception as e:
                    st.error(f"시스템 오류: {e}")

    if st.session_state.analyzed_data:
        data = st.session_state.analyzed_data
        excel_data = create_excel(paper_num, data['json'], data['images'], data['figs'], data['tbls'])

        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=excel_data,
            file_name=f"Analysis_v10.1_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
