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
st.title("📑 논문 분석 Pro [ver6.7 - Image Recovery]")
st.caption("✅ 이미지 추출 기능 복구 | 100px 이하 로고/아이콘 자동 삭제 | 한글 출력 필수")

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
# [5] 핵심 로직 함수 (안정성 강화)
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
        # 1. 텍스트 추출 및 캡션 위치 찾기
        text_blocks = page.get_text("blocks")
        for b in text_blocks:
            text = b[4].strip()
            final_text_content += text + "\n"

            # 캡션 인식 (Fig, Table, 그림, 표)
            if re.match(r"^(Fig|Figure|Table|그림|표)\s*\.?\s*\d+", text, re.IGNORECASE) and len(text) < 300:
                bbox = fitz.Rect(b[0], b[1], b[2], b[3])
                cap_type = "Table" if (text.startswith("Table") or text.startswith("표")) else "Figure"
                label_match = re.match(r"(Fig\.?|Figure|Table|그림|표)\s*\d+", text, re.IGNORECASE)
                label = label_match.group(0) if label_match else cap_type

                all_captions.append({
                    "page": page_num, "bbox": bbox, "text": text,
                    "type": cap_type, "label": label, "matched_img_id": None
                })

        # 2. 페이지 이미지 저장 (AI 텍스트 분석 보완용)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_data = Image.open(io.BytesIO(pix.tobytes("png")))
        all_page_images.append(img_data)

        # 3. 이미지 추출 (기존 방식 복구 + 필터링 강화)
        image_list = page.get_images(full=True)
        raw_rects = []

        for img in image_list:
            xref = img[0]
            # 이미지 위치 정보 가져오기
            try:
                img_rects = page.get_image_rects(xref)
                for r in img_rects:
                    # [핵심 필터] 너무 작은 이미지(로고, 아이콘)는 버림
                    if r.width < 100 or r.height < 100:
                        continue
                    raw_rects.append(r)
            except:
                continue

        # 겹치는 이미지 영역 병합
        merged_rects = merge_nearby_rectangles(raw_rects, distance=20)

        for rect in merged_rects:
            img_id = f"Image_{image_counter}"
            all_images_info.append({
                "id": img_id, "page": page_num, "bbox": rect, "matched_caption": None
            })
            image_counter += 1

    # 4. 캡션과 이미지 매칭 (거리 기반)
    for cap in all_captions:
        best_img = None
        min_score = float('inf')

        # 같은 페이지에 있는 이미지들만 후보로 선정
        candidates = [img for img in all_images_info if img["page"] == cap["page"] and img["matched_caption"] is None]

        for img in candidates:
            # Figure는 캡션이 보통 아래, Table은 위
            # 하지만 너무 엄격하면 놓칠 수 있으니 거리 점수(Score)제로 계산

            # 수직 거리
            if cap["type"] == "Figure":
                # 그림은 캡션보다 위에 있어야 함 (캡션 y0 - 이미지 y1)
                v_dist = (cap["bbox"].y0 - img["bbox"].y1)
            else:
                # 표는 캡션보다 아래에 있어야 함 (이미지 y0 - 캡션 y1)
                v_dist = (img["bbox"].y0 - cap["bbox"].y1)

            # 방향이 맞지 않으면 페널티 부여, 하지만 아주 가깝다면 허용
            if v_dist < -50:  # -50px 이상 반대 방향이면 제외
                continue

            abs_v_dist = abs(v_dist)

            # 수평 거리 (중앙 정렬 여부)
            cap_center_x = (cap["bbox"].x0 + cap["bbox"].x1) / 2
            img_center_x = (img["bbox"].x0 + img["bbox"].x1) / 2
            h_align_dist = abs(cap_center_x - img_center_x)

            # 점수 계산 (거리가 가까울수록 좋음)
            total_score = abs_v_dist + (h_align_dist * 0.5)

            if total_score < min_score:
                min_score = total_score
                best_img = img

        if best_img:
            cap["matched_img_id"] = best_img["id"]
            best_img["matched_caption"] = cap["label"]

    # 5. 최종 이미지 추출 및 저장
    extracted_images_map = {}
    for img_info in all_images_info:
        page = doc[img_info["page"]]
        rect = img_info["bbox"]

        # 이미지 영역 캡처
        padding = 10
        clip_rect = fitz.Rect(rect.x0 - padding, rect.y0 - padding, rect.x1 + padding, rect.y1 + padding) & page.rect

        try:
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat, clip=clip_rect)
            img_bytes = pix.tobytes("png")

            img_id = img_info["id"]
            initial_label = img_info["matched_caption"] if img_info["matched_caption"] else "Figure"

            extracted_images_map[img_id] = {
                "id": img_id, "page": img_info["page"] + 1, "bytes": img_bytes,
                "initial_label": initial_label, "real_label": initial_label
            }
        except:
            continue

    extracted_images = list(extracted_images_map.values())
    return final_text_content, extracted_images, all_page_images


def get_gemini_analysis(text, total_images, all_page_images):
    inputs = []

    # [프롬프트] 한국어 필수, 이미지 매칭 강조
    prompt = f"""
    너는 한국어 논문 분석 전문가야. 제공된 자료를 보고 JSON을 추출해.

    [절대 규칙]
    1. **모든 요약(Summary)은 반드시 '한국어(Korean)'로 작성해.**
    2. **요약은 '개조식(Bullet Points)'으로 작성해.**
    3. **이미지 매칭:**
       - `referenced_images` 리스트를 만들 때, 내가 제공한 이미지 리스트의 `real_label`(예: 그림 1)과 정확히 매칭해.
       - 만약 매칭되는 이미지가 없다면 억지로 넣지 마.

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

    # 텍스트 부족 시 이미지 전송
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
            with st.spinner(f"[{SELECTED_MODEL_NAME}] 분석 중... (이미지 복구 모드)"):
                try:
                    text, images, all_page_imgs = extract_data_from_pdf(uploaded_file)

                    if len(text.strip()) < 500:
                        st.warning("⚠️ 텍스트 추출이 불안정하여 전체 페이지 분석을 병행합니다.")
                    else:
                        st.info(f"✅ 텍스트 및 {len(images)}개의 유효 이미지(로고 제외) 추출 완료!")

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
            file_name=f"Analysis_v6.7_{paper_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
