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
st.title("📑 논문 분석 Pro [ver10.4 - Hybrid Summary]")
st.caption("✅ 스캔본(이미지 문서) 완벽 대응 | 텍스트 없으면 AI가 눈으로 보고 요약 | 이미지 짤림 방지")

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

        # Vision 성능이 좋은 모델 우선
        preferred = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']
        available_models.sort(key=lambda x: (x not in preferred, x))

        selected_model_name = st.selectbox(
            "✅ 모델 선택 (2.5-flash 기본)",
            available_models,
            index=0
        )
        SELECTED_MODEL_NAME = f"models/{selected_model_name}"
        st.success(f"연결됨: {selected_model_name}")

        if "pro" in selected_model_name:
            st.info("💡 Pro 모델: 스캔본 인식률이 더 높습니다.")
        else:
            st.info("⚡ Flash 모델: 속도가 빠릅니다.")

    except Exception as e:
        st.error(f"모델 목록 오류: {e}")
        st.stop()

model = genai.GenerativeModel(SELECTED_MODEL_NAME)


# -----------------------------------------------------------
# [4] 핵심 로직: AI Vision (좌표 추출)
# -----------------------------------------------------------
def detect_regions_with_gemini(page_image):
    prompt = """
    Look at this research paper page. 
    Detect all **Figures** and **Tables**.

    [Rules]
    1. Return Bounding Boxes in **normalized coordinates (0 to 1000)**: [ymin, xmin, ymax, xmax].
    2. **IMPORTANT: Be GENEROUS with the bounding box.** - Expand the box to include ALL labels, axis titles, legends, and the full caption text.
       - Do not cut off the edges.
    3. **ALWAYS return 4 numbers** for the box.
    4. Group multiple parts (a, b) into ONE box.
    5. Output JSON list.

    Example:
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

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_pages = len(doc)

    for page_num, page in enumerate(doc):
        status_text.text(f"🔍 AI가 {page_num + 1}/{total_pages} 페이지를 정밀 분석 중입니다...")
        progress_bar.progress((page_num + 1) / total_pages)

        # 텍스트 추출 (요약용)
        text_on_page = page.get_text()
        final_text_content += text_on_page + "\n"

        # 이미지 변환 (Vision 분석용)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data_bytes = pix.tobytes("png")
        pil_image = Image.open(io.BytesIO(img_data_bytes))
        all_page_images.append(pil_image)

        # AI Vision 좌표 요청
        detected_objects = detect_regions_with_gemini(pil_image)

        page_width = page.rect.width
        page_height = page.rect.height

        if detected_objects:
            for obj in detected_objects:
                label = obj.get("label", "Unknown")
                obj_type = obj.get("type", "Figure")
                box = obj.get("box_2d")

                if not box or not isinstance(box, list) or len(box) != 4:
                    continue

                ymin, xmin, ymax, xmax = box

                # 좌표 변환
                real_x0 = (xmin / 1000) * page_width
                real_y0 = (ymin / 1000) * page_height
                real_x1 = (xmax / 1000) * page_width
                real_y1 = (ymax / 1000) * page_height

                # [여백 확장 로직]
                pad_x = 20
                pad_y = 20

                final_x0 = max(0, real_x0 - pad_x)
                final_x1 = min(page_width, real_x1 + pad_x)

                # 캡션 방향 확장
                if "Figure" in obj_type or "Fig" in label:
                    final_y0 = max(0, real_y0 - pad_y)
                    final_y1 = min(page_height, real_y1 + 60)  # 아래로 60px
                elif "Table" in obj_type or "Tab" in label:
                    final_y0 = max(0, real_y0 - 60)  # 위로 60px
                    final_y1 = min(page_height, real_y1 + pad_y)
                else:
                    final_y0 = max(0, real_y0 - pad_y)
                    final_y1 = min(page_height, real_y1 + pad_y)

                crop_rect = fitz.Rect(final_x0, final_y0, final_x1, final_y1)

                if crop_rect.width < 50 or crop_rect.height < 50: continue

                try:
                    clip_pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=crop_rect)
                    img_bytes = clip_pix.tobytes("png")

                    img_id = f"Image_{image_counter}"
                    image_counter += 1

                    extracted_images_map[img_id] = {
                        "id": img_id,
                        "page": page_num + 1,
                        "bytes": img_bytes,
                        "initial_label": label,
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
    # [프롬프트] 요약 요청
    prompt = f"""
    너는 논문 분석 전문가야. 제공된 자료를 보고 내용을 한국어로 요약해.

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

    # [핵심 로직] 텍스트가 충분한지 확인
    if len(text.strip()) > 500:
        # 텍스트 PDF: 텍스트로 분석 (빠름)
        inputs.append(f"[Text Data]:\n{text[:50000]}")
    else:
        # 스캔본 PDF: 이미지로 분석 (Vision)
        # 중요: 모든 페이지를 다 보내면 토큰 초과될 수 있으니 최대 20페이지만 전송
        inputs.append("⚠️ 텍스트 데이터가 부족합니다(스캔 문서). 페이지 이미지를 보고 내용을 분석하세요.")

        max_pages_to_send = 20
        target_images = all_page_images[:max_pages_to_send]

        for i, img in enumerate(target_images):
            inputs.append(f"Page {i + 1} Image:")
            inputs.append(img)

    try:
        response = model.generate_content(inputs, generation_config={"response_mime_type": "application/json"})
        return json.loads(response.text)
    except Exception as e:
        return {"error": str(e)}


# -----------------------------------------------------------
# [6] 엑셀 생성 및 유틸리티
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
        if st.session_state.analyzed_data and st.session_state.analyzed_data['file_name'] == uploaded_file.name:
            st.success("⚡ 저장된 분석 결과를 불러옵니다.")
        else:
            with st.spinner(f"[{SELECTED_MODEL_NAME}] AI Vision 분석 중..."):
                try:
                    # 1. 이미지 추출
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if len(text.strip()) < 500:
                        st.warning("⚠️ 텍스트 데이터가 부족합니다(스캔 문서). 이미지 기반 분석을 수행합니다.")

                    if not images:
                        st.warning("⚠️ AI가 그림/표를 찾지 못했습니다. 모델을 '1.5-pro'로 변경해보세요.")
                    else:
                        st.info(f"✅ AI가 {len(images)}개의 그림/표 영역을 인식했습니다!")

                    # 2. 내용 분석 (텍스트 or 이미지)
                    result = get_gemini_analysis(text, len(images), all_page_imgs)

                    if "error" in result:
                        st.error(f"AI 분석 오류: {result['error']}")
                    else:
                        # 3. 매칭 및 정렬
                        ref_imgs = result.get('referenced_images', [])
                        final_figs, final_tbls = [], []

                        for img in images:
                            img_label = img['initial_label']
                            matched_caption = "설명 없음"
                            for ref in ref_imgs:
                                ref_l = standardize_label_to_korean(ref.get('real_label', ''))[2]
                                img_l = standardize_label_to_korean(img_label)[2]
                                if ref_l == img_l:
                                    matched_caption = ref.get('caption', '-')
                                    break

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
            file_name=f"Analysis_v10.4_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
