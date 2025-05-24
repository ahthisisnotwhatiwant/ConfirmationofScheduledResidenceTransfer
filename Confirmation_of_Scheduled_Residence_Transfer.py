import streamlit as st
from datetime import date
import os
import uuid
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path, convert_from_bytes
from io import BytesIO
import textwrap
from streamlit_drawable_canvas import st_canvas
import base64
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import formataddr
import re

# 경로 설정 (템플릿 PDF)
PDF_TEMPLATE_PATH = "consent.pdf"
TRANSFER_FORM_PATH = "transfer.pdf"
FONT_PATH = "malgun.ttf"
CONSENT_SAMPLE_PATH = "consent_sample.pdf"
TRANSFER_SAMPLE_PATH = "transfer_sample.pdf"
XLSX_FILE_PATH = "school_data.xlsx"

# 환경 변수에서 이메일 설정 정보 읽어오기
MAIL_FROM = os.getenv("MAIL_FROM")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))

st.set_page_config(page_title="전입예정확인서", layout="centered")

# 학년을 영어 형식으로 변환하는 함수
def grade_to_english(grade):
    number = re.search(r'\d+', grade)
    if number:
        return f"{number.group()}gr"
    return grade

# PDF 파일을 이미지로 변환하는 함수
def convert_pdf_to_images(pdf_path, dpi=150):
    try:
        images = convert_from_path(pdf_path, dpi=dpi)
        return images
    except Exception as e:
        st.error(f"PDF를 이미지로 변환 중 오류 발생: {e}")
        return None

# 기존 CSS 유지
st.markdown("""
    <style>
    .title {
        font-size: 2.5rem;
        font-weight: bold;
        color: #4c51bf;
        text-align: center;
        padding-bottom: 1rem;
        margin-bottom: 2rem;
        background: linear-gradient(to right, #f0f2ff, #ffffff);
        -webkit-background-clip: text;
        color: transparent;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    .pdf-viewer {
        width: 100%;
        height: 500px;
        border: 1px solid #d1d5db;
        margin-bottom: 2rem;
    }
    .instruction-message {
        background-color: #f0fdf4;
        color: #15803d;
        padding: 0.75rem;
        margin-bottom: 1rem;
        border-radius: 0.375rem;
        border: 1px solid #bbf7d0;
        font-size: 0.875rem;
        text-align: center;
    }
    </style>
    <h1 class="title">전입예정확인서</h1>
""", unsafe_allow_html=True)

# 사용자 안내
st.markdown('<div class="instruction-message">----------  목  적  ---------- <br> 신설학교 학급 편성을 위한 정보 수집<br>----------  순  서  ---------- <br> ①지역 및 학교 → ②개인정보 수집·이용 동의서 → ③전입예정확인서 → ④제출</div>', unsafe_allow_html=True)

# Streamlit Session State 초기화
if 'stage' not in st.session_state:
    st.session_state.stage = 1
    st.session_state.agree_to_collection = "none"
    st.session_state.schools_by_region = {}
    st.session_state.selected_region = ""
    st.session_state.selected_school = ""
    st.session_state.student_name = ""
    st.session_state.move_date = None
    st.session_state.pdf_bytes = None
    st.session_state.filename = None

# 입력 검증 함수
def validate_inputs(student_name, parent_name, student_phone, parent_phone, address, next_grade):
    if not all([student_name, parent_name, student_phone, parent_phone, address, next_grade]):
        return False, "모든 필드를 입력하세요."
    phone_pattern = r'^\d{3}-\d{4}-\d{4}$'
    if not (re.match(phone_pattern, student_phone) and re.match(phone_pattern, parent_phone)):
        return False, "전화번호 형식이 올바르지 않습니다 → 옳은 예: 010-0000-0000"
    return True, ""

# 이메일 발송 함수
def send_pdf_email(pdf_data, filename, recipient_email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if not re.match(pattern, recipient_email):
        st.error(f"유효하지 않은 이메일 주소입니다: {recipient_email}")
        return False

    parts = filename.split('_')
    if len(parts) >= 3:
        grade = parts[2].replace('.pdf', '')
        english_grade = grade_to_english(grade)
        email_filename = f"Confirmation.of.Scheduled.Residence.Transfer_{english_grade}.pdf"
    else:
        email_filename = "Confirmation.of.Scheduled.Residence.Transfer.pdf"

    msg = MIMEMultipart()
    msg['From'] = formataddr((str(Header("전입예정확인서 시스템", 'utf-8')), MAIL_FROM))
    msg['To'] = recipient_email
    msg['Subject'] = f"전입예정확인서({filename})"

    body = f"안녕하세요.\n\n{filename}가 제출되었습니다.\n첨부된 PDF 파일을 저장 후 이상이 없는지 확인하여 주세요.\n편리한 관리를 위해 파일명 변경을 권장드립니다.\n\n감사합니다."
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    part = MIMEBase('application', 'pdf')
    part.set_payload(pdf_data)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{email_filename}"', filename=('utf-8', '', email_filename))
    part.add_header('Content-Type', f'application/pdf; name="{email_filename}"')
    msg.attach(part)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(MAIL_FROM, MAIL_PASSWORD)
        server.sendmail(MAIL_FROM, recipient_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"이메일 발송 실패: {e}")
        st.error("이메일 설정을 확인하고 다시 시도해주세요.")
        return False

# 1단계: 지역 및 학교 선택
if st.session_state.stage == 1:
    st.subheader("1단계: 지역 및 학교")
    st.markdown('<div class="instruction-message">전입 예정 지역 및 전학 예정 학교를 선택하세요.</div>', unsafe_allow_html=True)

    try:
        df = pd.read_excel(XLSX_FILE_PATH)
        if not all(col in df.columns for col in ['지역', '학교', '이메일']):
            st.error("XLSX 파일에 '지역', '학교', '이메일' 컬럼이 있어야 합니다. 파일 내용을 확인하고 다시 시도해주세요.")
            st.stop()
        st.session_state.schools_by_region = df.groupby('지역')['학교'].apply(list).to_dict()
        regions = list(st.session_state.schools_by_region.keys())
    except Exception as e:
        st.error(f"XLSX 파일을 읽는 중 오류가 발생했습니다: {e}. 파일 경로 및 형식을 확인해주세요. 경로: {XLSX_FILE_PATH}")
        st.stop()

    st.session_state.selected_region = st.selectbox("전입 예정 지역을 선택하세요.", regions)

    available_schools = st.session_state.schools_by_region.get(st.session_state.selected_region, [])
    if not available_schools:
        st.warning("선택한 지역에 학교 정보가 없습니다. 다른 지역을 선택해주세요.")
        st.session_state.selected_school = ""
    else:
        st.session_state.selected_school = st.selectbox("전학 예정 학교를 선택하세요.", available_schools)

    if st.button("✒️다음 단계로"):
        if st.session_state.selected_region and st.session_state.selected_school:
            st.session_state.stage = 2
            st.rerun()
        else:
            st.warning("지역과 학교를 모두 선택하세요.")

# 2단계: 개인정보 수집·이용 동의서
elif st.session_state.stage == 2:
    st.subheader("2단계: 개인정보 수집·이용 동의서")
    st.markdown('<div class="instruction-message">개인정보 수집·이용 동의서를 확인 후 진행하세요.</div>', unsafe_allow_html=True)

    # 샘플 PDF를 이미지로 표시
    consent_images = convert_pdf_to_images(CONSENT_SAMPLE_PATH, dpi=150)
    if consent_images:
        with st.expander("📄 개인정보 수집·이용 동의서 예시", expanded=True):
            for i, image in enumerate(consent_images):
                st.image(image, use_container_width=True)
    else:
        st.error("동의서 샘플 PDF를 불러올 수 없습니다. 파일 경로를 확인해주세요.")

    st.markdown("☞ 위와 같이 개인정보 수집·이용에 동의하십니까?")
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.agree_to_collection = st.checkbox("동의합니다.")
    with col2:
        st.session_state.disagree_to_collection = st.checkbox("동의하지 않습니다.")
    if st.session_state.agree_to_collection and st.session_state.disagree_to_collection:
        st.warning("'동의합니다.'와 '동의하지 않습니다.' 중 **하나**만 선택하세요.")
        st.session_state.agree_to_collection = False
        st.session_state.disagree_to_collection = False
    if st.session_state.agree_to_collection:
        if st.button("✒️다음 단계로"):
            st.session_state.stage = 3
            st.rerun()
    elif st.session_state.disagree_to_collection:
        st.warning("개인정보 수집·이용에 동의 시에만 다음 단계로 진행할 수 있습니다.")

# 3단계: 전입예정확인서
elif st.session_state.stage == 3:
    st.subheader("3단계: 전입예정확인서")
    st.markdown('<div class="instruction-message">작성란 예시를 지운 후 작성하세요.</div>', unsafe_allow_html=True)

    # 샘플 PDF를 이미지로 표시
    transfer_images = convert_pdf_to_images(TRANSFER_SAMPLE_PATH, dpi=150)
    if transfer_images:
        with st.expander("📄 전입예정확인서 예시", expanded=True):
            for i, image in enumerate(transfer_images):
                st.image(image, use_container_width=True)
    else:
        st.error("전입예정확인서 샘플 PDF를 불러올 수 없습니다. 파일 경로를 확인해주세요.")

    col1, col2 = st.columns(2)
    with col1:
        st.session_state.student_name = st.text_input("학생 이름", value="000")
        student_school = st.text_input("현 소속 학교 및 학년", value="00초등학교 0학년")
        student_phone = st.text_input("학생 휴대전화 번호", value="010-0000-0000")
        st.session_state.move_date = st.date_input("전입 예정일", value=date.today())
        school_name = st.text_input("전학 예정 학교", value=st.session_state.selected_school)
    with col2:
        parent_name = st.text_input("법정대리인 이름", value="000")
        relationship = st.text_input("학생과의 관계", value="부, 모 등")
        parent_phone = st.text_input("법정대리인 휴대전화 번호", value="010-0000-0000")
        address = st.text_input("전입 예정 주소", value="00택지 A-0블록 00아파트 00동 00호")
        next_grade = st.text_input("전학 예정 학년", value="0학년")

    col1, col2 = st.columns(2)
    with col1:
        st.write("학생 서명")
        canvas_student = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=5,
            background_color="rgba(255, 255, 255, 0)",
            height=150,
            width=300,
            drawing_mode="freedraw",
            key="student_sign_canvas"
        )
    with col2:
        st.write("법정대리인 서명")
        canvas_parent = st_canvas(
            fill_color="rgba(255, 255, 255, 0)",
            stroke_width=5,
            background_color="rgba(255, 255, 255, 0)",
            height=150,
            width=300,
            drawing_mode="freedraw",
            key="parent_sign_canvas"
        )

    if st.button("✒️다음 단계로"):
        valid, error = validate_inputs(st.session_state.student_name, parent_name, student_phone, parent_phone, address, next_grade)
        if not valid:
            st.error(error)
            st.stop()
        try:
            # 서명 비율 체크
            def calculate_signature_coverage(image_data):
                alpha_channel = image_data[:, :, 3]
                drawn_pixels = (alpha_channel > 0).sum()
                total_pixels = image_data.shape[0] * image_data.shape[1]
                return drawn_pixels / total_pixels

            student_coverage = calculate_signature_coverage(canvas_student.image_data)
            parent_coverage = calculate_signature_coverage(canvas_parent.image_data)

            if student_coverage < 0.05 or parent_coverage < 0.05:
                st.warning("학생과 법정대리인 모두 올바르게 서명하세요.")
                st.stop()

            student_sign_path = f"student_sign_{uuid.uuid4()}.png"
            parent_sign_path = f"parent_sign_{uuid.uuid4()}.png"
            Image.fromarray(canvas_student.image_data.astype('uint8'), mode='RGBA').save(student_sign_path, optimize=True)
            Image.fromarray(canvas_parent.image_data.astype('uint8'), mode='RGBA').save(parent_sign_path, optimize=True)

            pages1 = convert_from_path(PDF_TEMPLATE_PATH, dpi=200)
            page1 = pages1[0].convert('RGBA')
            pages2 = convert_from_path(TRANSFER_FORM_PATH, dpi=200)
            page2 = pages2[0].convert('RGBA')
            draw1 = ImageDraw.Draw(page1)
            draw2 = ImageDraw.Draw(page2)

            consent_positions = {
                "{{date.today}}": [(975, 1540)],
                "{{student_name}}": [(815, 1685)],
                "{{student_sign_path}}": [(1050, 1655)],
                "{{parent_name}}": [(815, 1810)],
                "{{parent_sign_path}}": [(1050, 1800)],
                "{{school_name}}": [(937, 1982)],
            }
            transfer_positions = {
                "{{student_name}}": [(480, 457), (815, 1760)],
                "{{parent_name}}": [(1140, 457), (815, 1885)],
                "{{student_school}}": [(440, 555)],
                "{{relationship}}": [(1140, 555)],
                "{{student_phone}}": [(462, 650)],
                "{{parent_phone}}": [(1105, 650)],
                "{{move_date}}": [(462, 847)],
                "{{address}}": [(1140, 829), (520, 1185)],
                "{{school_name}}": [(462, 1048), (320, 1245), (937, 2057)],
                "{{next_grade}}": [(1115, 1048), (920, 1245)],
                "{{date.today}}": [(975, 1610)],
                "{{student_sign_path}}": [(1050, 1740)],
                "{{parent_sign_path}}": [(1050, 1880)],
            }

            def get_font(key, idx):
                if key == "{{address}}" and idx == 0:
                    return ImageFont.truetype(FONT_PATH, 25)
                if key == "{{address}}" and idx == 1:
                    return ImageFont.truetype(FONT_PATH, 34)
                return ImageFont.truetype(FONT_PATH, 42)

            consent_map = {
                "{{student_name}}": st.session_state.student_name,
                "{{parent_name}}": parent_name,
                "{{date.today}}": date.today().strftime("%Y년 %m월 %d일"),
                "{{school_name}}": school_name,
            }
            transfer_map = {
                **consent_map,
                "{{student_school}}": student_school,
                "{{relationship}}": relationship,
                "{{student_phone}}": student_phone,
                "{{parent_phone}}": parent_phone,
                "{{move_date}}": st.session_state.move_date.strftime("%Y년 %m월 %d일"),
                "{{address}}": address,
                "{{next_grade}}": next_grade,
            }

            def draw_texts(draw, positions, data_map, is_transfer=False):
                for key, coords in positions.items():
                    for idx, (x, y) in enumerate(coords):
                        text = data_map.get(key, "")
                        font = get_font(key, idx)
                        if not is_transfer:
                            if key in ["{{student_name}}", "{{parent_name}}", "{{student_sign_path}}", "{{parent_sign_path}}"]:
                                x -= 15
                        else:
                            if key == "{{address}}":
                                if idx == 0:
                                    x -= 7
                                    text = "\n".join(textwrap.wrap(text, width=10))
                                elif idx == 1:
                                    x -= 50
                            if key == "{{next_grade}}" and idx == 1:
                                x += 50
                        draw.text((x, y), text, font=font, fill='black')

            draw_texts(draw1, consent_positions, consent_map, is_transfer=False)
            sign1 = Image.open(student_sign_path).resize((312, 104)).convert('RGBA')
            sign2 = Image.open(parent_sign_path).resize((312, 104)).convert('RGBA')
            for x, y in consent_positions.get("{{student_sign_path}}", []):
                page1.paste(sign1, (x - 15, y), sign1)
            for x, y in consent_positions.get("{{parent_sign_path}}", []):
                page1.paste(sign2, (x - 15, y), sign2)

            draw_texts(draw2, transfer_positions, transfer_map, is_transfer=True)
            for x, y in transfer_positions.get("{{student_sign_path}}", []):
                page2.paste(sign1, (x, y), sign1)
            for x, y in transfer_positions.get("{{parent_sign_path}}", []):
                page2.paste(sign2, (x, y), sign2)

            buffer = BytesIO()
            page1 = page1.convert('RGB')
            page2 = page2.convert('RGB')
            page1.save(buffer, format='PDF', quality=70)
            page2.save(buffer, format='PDF', append=True, save_all=True, quality=70)
            pdf_bytes = buffer.getvalue()
            filename = f"전입예정확인서_{school_name}_{next_grade}.pdf"

            st.session_state.pdf_bytes = pdf_bytes
            st.session_state.filename = filename
            st.session_state.stage = 4
            st.rerun()

        except Exception as e:
            st.error(f"PDF 생성 중 오류 발생: {e}")
        finally:
            try:
                if 'student_sign_path' in locals() and os.path.exists(student_sign_path):
                    os.remove(student_sign_path)
                if 'parent_sign_path' in locals() and os.path.exists(parent_sign_path):
                    os.remove(parent_sign_path)
            except Exception as e:
                st.warning(f"임시 파일 삭제 중 오류 발생: {e}")

# 4단계: 미리보기 및 제출
elif st.session_state.stage == 4:
    st.subheader("4단계: 미리보기 및 제출")
    st.markdown('<div class="instruction-message">미리보기를 통해 최종 확인 후 제출하세요.</div>', unsafe_allow_html=True)

    if st.session_state.pdf_bytes and st.session_state.filename:
        try:
            # PDF를 이미지로 변환
            from pdf2image import convert_from_bytes
            images = convert_from_bytes(st.session_state.pdf_bytes, dpi=150)

            # 이미지 미리보기를 확장 가능한 섹션에 표시
            with st.expander("📄 전입예정확인서 미리보기", expanded=True):
                for i, image in enumerate(images):
                    st.image(image, use_container_width=True)

            # PDF 다운로드 버튼
            st.download_button(
                label="💾 전입예정확인서 내려받기",
                data=st.session_state.pdf_bytes,
                file_name=st.session_state.filename,
                mime='application/pdf'
            )

            if st.button("📮 전입예정확인서 제출하기"):
                with st.spinner("제출 중입니다. 잠시만 기다려 주세요."):
                    try:
                        df = pd.read_excel(XLSX_FILE_PATH)
                        email_series = df[df['학교'] == st.session_state.selected_school]['이메일']
                        if email_series.empty:
                            st.error(f"학교 '{st.session_state.selected_school}'에 해당하는 이메일이 없습니다.")
                            st.error("오류가 발생했습니다. 다시 처음부터 진행해주세요.")
                            st.stop()
                        selected_school_email = email_series.values[0]
                        if send_pdf_email(st.session_state.pdf_bytes, st.session_state.filename, selected_school_email):
                            st.success("정상적으로 제출되었습니다. 협조해 주셔서 감사합니다.")
                        else:
                            st.error("오류가 발생했습니다. 다시 처음부터 진행해주세요.")
                    except Exception as e:
                        st.error(f"이메일 발송 중 오류 발생: {e}")
                        st.error("오류가 발생했습니다. 다시 처음부터 진행해주세요.")
        except Exception as e:
            st.error(f"PDF 미리보기 이미지 생성 중 오류 발생: {e}")
            st.error("PDF 파일을 다운로드하여 확인해 주세요.")
            st.download_button(
                label="💾 전입예정확인서 내려받기",
                data=st.session_state.pdf_bytes,
                file_name=st.session_state.filename,
                mime='application/pdf'
            )
    else:
        st.error("PDF가 생성되지 않았습니다. 3단계로 돌아가 PDF를 생성해 주세요.")
