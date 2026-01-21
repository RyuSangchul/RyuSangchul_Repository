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
st.title("📑 논문 분석 Pro [ver6.4 - Classification Fix]")
st.caption("✅ Figure/Table 구분 강화 | 번호(1, 2, 3) 자동 정렬 기능 추가")

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
# [4] 유틸리티 함수 (강화됨)
# -----------------------------------------------------------
def normalize_id(ref_text):
    """Image_1 -> Image_1 변환"""
    nums = re.findall(r'\d+', str(ref_text))
    return f"Image_{nums[0]}" if nums else None


def standardize_label(label_text):
    """
    [핵심 기능] 라벨 텍스트를 분석해서 (타입, 번호) 튜플을 반환합니다.
    예: "Fig. 3 결과 그래프" -> ("Figure", 3, "Fig. 3 결과 그래프")
    예: "Table 2. 실험 조건" -> ("Table", 2, "Table 2. 실험 조건")
    """
    if not label_text:
        return ("Unknown", 999, "Unknown")

    label_upper = str(label_text).upper()

    # 1. 타입 결정
    detected_type = "Figure"  # 기본값
    if "TAB" in label_upper or "표" in label_upper:
        detected_type = "Table"
    elif "FIG" in label_upper or "그림" in label_upper:
        detected_type = "Figure"

    # 2. 번호 추출 (숫자 찾기)
    # "Figure 1", "Fig.1", "Table-3" 등에서 숫자만 뽑음
    nums = re.findall(r'\d+', label_text)
    if nums:
        detected_num = int(nums[0])
    else:
        detected_num = 999  # 숫자가 없으면 맨 뒤로 보냄

    return (detected_type, detected_num, label_text)


def merge_nearby_rectangles(rects, distance=20):
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

    all_page_images = []
    all_captions = []
    all_images_info = []

    for page_num, page in enumerate(doc):
        text_blocks = page.get_text("blocks")
        for b in text_blocks:
            text = b[4].strip()
            final_text_content += text + "\n"

            # 캡션 후보 식별 (Fig 또는 Table로 시작하는 짧은 문장)
            if (text.startswith("Fig") or text.startswith("Table") or text.startswith("그림") or text.startswith(
                    "표")) and len(text) < 500:
                bbox = fitz.Rect(b[0], b[1], b[2], b[3])

                cap_type = "Table" if (text.startswith("Table") or text.startswith("표")) else "Figure"

                # 라벨 추출 (예: Fig. 1)
                label_match = re.match(r"(Fig\.?|Figure|Table|그림|표)\s*\d+", text, re.IGNORECASE)
                label = label_match.group(0) if label_match else cap_type

                all_captions.append({
                    "page": page_num, "bbox": bbox, "text": text,
                    "type": cap_type, "label": label, "matched_img_id": None
                })

        # 페이지 이미지 저장 (AI 전송용)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_data = Image.open(io.BytesIO(pix.tobytes("png")))
        all_page_images.append(img_data)

        # 논문 내부 이미지 좌표 추출
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

    # 캡션 매칭 (위치 기반)
    for cap in all_captions:
        best_img = None
        min_score = float('inf')
        candidates = [img for img in all_images_info if img["page"] == cap["page"] and img["matched_caption"] is None]

        for img in candidates:
            # Figure는 보통 캡션이 아래, Table은 보통 캡션이 위에 있음 (일반적 규칙)
            # 하지만 엄격하게 제한하지 않고 거리 점수로 판단
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

    # 최종 이미지 맵 생성
    extracted_images_map = {}
    for img_info in all_images_info:
        page = doc[img_info["page"]]
        rect = img_info["bbox"]
        padding = 35
        clip_rect = fitz.Rect(rect.x0 - padding, rect.y0 - padding, rect.x1 + padding, rect.y1 + padding) & page.rect
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat, clip=clip_rect)
        img_bytes = pix.tobytes("png")

        img_id = img_info["id"]
        # 캡션 매칭된 것이 있으면 그것을 우선 사용
        initial_label = img_info["matched_caption"] if img_info["matched_caption"] else "Unknown"

        extracted_images_map[img_id] = {
            "id": img_id, "page": img_info["page"] + 1, "bytes": img_bytes,
            "initial_label": initial_label, "real_label": initial_label
        }

    extracted_images = list(extracted_images_map.values())
    return final_text_content, extracted_images, all_page_images


def get_gemini_analysis(text, total_images, all_page_images):
    inputs = []

    # [프롬프트 강화] 분류 및 번호 인식을 강력하게 지시
    prompt = f"""
    너는 논문 분석 전문가야. 제공된 자료를 보고 JSON을 추출해.

    [핵심 지시사항]
    1. **요약(summary)은 반드시 '개조식(Bullet Points)'으로 작성.**
    2. **이미지 분류(Classification)를 정확히 수행해.**
       - 이미지 내용을 보고 **'Figure(그림/그래프)'**인지 **'Table(표)'**인지 판단해.
       - 텍스트에 "Table 1", "Figure 3" 같은 번호가 있다면 그 번호를 반드시 `real_label`에 적어.
       - 예시: `real_label`: "Figure 1", `real_label`: "Table 2"
    3. 텍스트가 깨져 보이면 제공된 '페이지 이미지'를 읽어서 내용을 파악해.

    [요청 항목]
    0. title, author, affiliation, year, purpose
    1. 요약 (intro_summary, body_summary, conclusion_summary)
    2. key_images_desc
    3. referenced_images (여기에 각 이미지의 ID와 정확한 라벨을 적어줘)

    [출력 포맷 JSON]
    {{
        "title": "...",
        "author": "...", "affiliation": "...", "year": "...", "purpose": "...",
        "intro_summary": "- ...", 
        "body_summary": "- ...", 
        "conclusion_summary": "- ...",
        "key_images_desc": "...",
        "referenced_images": [ 
            {{ "img_id": "Image_1", "real_label": "Figure 1", "caption": "설명" }},
            {{ "img_id": "Image_2", "real_label": "Table 1", "caption": "설명" }}
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
        if len(all_page_images) > max_pages:
            inputs.append("[System: 뒷부분 페이지 일부 생략됨]")

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
        # 리스트 에러 방지
        if isinstance(content, list):
            content = "\n".join(map(str, content))
        elif content is None:
            content = "-"

        ws1.write(current_row, 0, label, header_style)
        ws1.write(current_row, 1, str(content), content_style)
        current_row += 1

    # Figure 섹션 쓰기
    if final_figures:
        current_row += 1
        ws1.write(current_row, 0, "Figures (그림)", header_style)
        ws1.write(current_row, 1, "▼ 주요 그림 목록", header_style)
        current_row += 1
        if current_row % 2 != 0: current_row += 1
        for item in final_figures:
            _write_row_dynamic(ws1, item, images, current_row, fig_style, content_style)
            current_row += 2

    # Table 섹션 쓰기
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

    label_text = item.get('real_label', 'Figure')
    caption_text = item.get('caption', '설명 없음')

    ws.write(row, 0, label_text, label_fmt)
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
            with st.spinner(f"[{SELECTED_MODEL_NAME}] 분석 중... (정밀 분류 모드)"):
                try:
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if len(text.strip()) < 500:
                        st.warning(f"⚠️ 텍스트 추출 실패! 논문 전체({len(all_page_imgs)}페이지)를 이미지로 읽습니다.")
                    else:
                        st.info("✅ 텍스트 추출 성공! 빠른 분석 모드로 실행합니다.")

                    result = get_gemini_analysis(text, len(images), all_page_imgs)

                    if "error" in result:
                        st.error(f"AI 분석 오류: {result['error']}")
                    else:
                        ref_imgs = result.get('referenced_images', [])
                        final_figs, final_tbls = [], []

                        # [분류 로직 강화]
                        for item in ref_imgs:
                            raw_label = item.get('real_label', 'Unknown')

                            # standardize_label 함수를 통해 (타입, 번호, 원본라벨) 분리
                            detected_type, detected_num, _ = standardize_label(raw_label)

                            # 정렬을 위해 번호 정보를 item에 추가
                            item['sort_num'] = detected_num

                            if detected_type == "Table":
                                final_tbls.append(item)
                            else:
                                final_figs.append(item)

                        # 정렬: 번호 순서대로 (1, 2, 3...)
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
            file_name=f"Analysis_v6.4_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
