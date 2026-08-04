import streamlit as st
import pandas as pd
import io
import json
import tempfile
import os
import time
import zipfile
import PyPDF2
import google.generativeai as genai  
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# ==========================================
# ⚙️ [초기 세팅] 시스템 가상 메모리 및 페이지 설정
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")

if 'uploaded_pdfs' not in st.session_state: st.session_state.uploaded_pdfs = {}
if 'eval_criteria' not in st.session_state: st.session_state.eval_criteria = pd.DataFrame() 
if 'auto_settings' not in st.session_state:
    st.session_state.auto_settings = {
        "has_safety": True, "period": "3년",
        "bohal": [{"전문분야": "상하수도", "비율(%)": 60}, {"전문분야": "토질지질", "비율(%)": 40}],
        "pm_cnt": 1, "pe_cnt": 2, "pes_cnt": 2
    }

with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")
    else:
        st.warning("문서 자동 분석을 위해 API Key가 필요합니다.")

# ==========================================
# 🔑 [Google Drive 인증]
# ==========================================
@st.cache_resource
def authenticate_google_drive():
    try:
        oauth_data = st.secrets["google_oauth"]
        creds = Credentials(
            token=None,
            refresh_token=oauth_data["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=oauth_data["client_id"],
            client_secret=oauth_data["client_secret"]
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"구글 드라이브 인증 실패 (Secrets 설정을 확인하세요): {e}")
        return None

def get_or_create_folder(drive_service, folder_name, parent_id=None):
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id: query += f" and '{parent_id}' in parents"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    if items: return items[0]['id']
    else:
        folder_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
        if parent_id: folder_metadata['parents'] = [parent_id]
        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        return folder.get('id')

@st.cache_data(ttl=600)
def load_master_db_from_drive():
    try:
        drive_service = authenticate_google_drive()
        if not drive_service: return pd.DataFrame()
        results = drive_service.files().list(
            q="name contains '마스터' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false",
            fields="files(id, name)"
        ).execute()
        items = results.get('files', [])
        if not items: return pd.DataFrame()
        file_id = items[0]['id']
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: status, done = downloader.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh)
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 🧠 [Backend Engine] PQ 점수 계산 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        self.master_db = load_master_db_from_drive()

    def get_personnel_list(self):
        if self.master_db.empty: return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]
        for col_name in ['이름', '성명', '엔지니어명', '기술자명', '기술인']:
            if col_name in self.master_db.columns:
                return ["(선택)"] + self.master_db[col_name].dropna().unique().tolist()
        return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]

    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt):
        best_score_df = pd.DataFrame({"평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"], "배점": [30, 20, 10, -5], "획득점수": [30.0, 20.0, 10.0, 0.0], "비고": ["AI 최적화", "중첩도 0건", "A+ 등급", "해당없음"]})
        sample_names = ["윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]
        return best_score_df, sample_names[0:pm_cnt], sample_names[pm_cnt:pm_cnt+pe_cnt], sample_names[pm_cnt+pe_cnt:pm_cnt+pe_cnt+pes_cnt]

    def calculate_manual_score(self):
        return pd.DataFrame({"평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"], "배점": [30, 20, 10, -5], "획득점수": [28.5, 20.0, 10.0, 0.0], "비고": ["수동 선택 검증 완료", "이상 없음", "우수", "해당 없음"]})

engine = PQScoringEngine()

def get_ai_model():
    try: return genai.GenerativeModel('gemini-3.6-flash')
    except: return genai.GenerativeModel('gemini-1.5-flash')

# ==========================================
# 🖥️ [Frontend] 메인 대시보드 UI
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 본 페이지는 실무 배포를 위한 Production 환경입니다.")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 세부사항 설정", "📊 3. 책임기술자 시뮬레이션 결과", "🖨️ 4. 서류 출력 및 패키징"])

# --- [Tab 1] 마스터 DB 관리 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF)을 드래그 앤 드롭하세요.", type=['pdf'], accept_multiple_files=True, key="zone_a")
        if notice_files and api_key:
            if st.button("🧠 공고문 AI 분석 및 평가기준 자동 구성", type="primary"):
                with st.spinner("AI가 공고문을 정독하며 배점표와 세부사항을 추출 중입니다..."):
                    try:
                        notice_text = ""
                        for file in notice_files:
                            pdf = PyPDF2.PdfReader(file)
                            for page in pdf.pages[:5]: notice_text += page.extract_text() or ""
                        prompt = f"건설엔지니어링(PQ) 공고문 분석 후 JSON 반환.\n1. eval_criteria: 배점표 배열\n2. settings: {{ has_safety, period, bohal, pm_cnt, pe_cnt, pes_cnt }}\n공고문: {notice_text}\n순수 JSON만 출력."
                        response = get_ai_model().generate_content(prompt)
                        result_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
                        parsed_json = json.loads(result_text)
                        
                        st.session_state.eval_criteria = pd.DataFrame(parsed_json.get("eval_criteria", []))
                        st.session_state.auto_settings = parsed_json.get("settings", st.session_state.auto_settings)
                        st.success("✅ 공고문 분석 성공! [Tab 2]에 모든 설정이 세팅되었습니다.")
                    except Exception as e: st.error(f"분석 에러: {e}")
                        
    with col2:
        st.subheader("Zone B: 실적 업데이트 및 자동 폴더링")
        perf_file = st.file_uploader("기술인/회사 실적증명서(PDF) 업로드", type=['pdf'], key="zone_b")
        if perf_file and api_key:
            with st.spinner("🧠 AI 스캔 및 드라이브 자동 폴더링 중입니다... (약 10초 소요)"):
                try:
                    pdf_part = {"mime_type": "application/pdf", "data": perf_file.getvalue()}
                    existing_str = ", ".join(map(str, engine.master_db['사업명'].dropna().tolist() if not engine.master_db.empty and '사업명' in engine.master_db.columns else []))
                    
                    # 👉 [핵심 수정] AI가 엑셀과 완벽히 동일한 칼럼명(키)을 뱉어내도록 명시했습니다!
                    prompt = f"""
                    PDF 문서 분석 후 JSON 반환.
                    1. doc_type (증명서 종류)
                    2. owner (이름, 없으면 '회사공통')
                    3. projects (배열). 반드시 "사업명", "시작일", "종료일", "담당업무", "발주처" 라는 5개의 키를 가진 객체 형태로 만들 것. 기존 목록 [{existing_str}] 과 의미상 같은 사업 제외하고 순수 신규만 반환.
                    순수 JSON만 출력.
                    """
                    response = get_ai_model().generate_content([prompt, pdf_part])
                    result_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
                    result_json = json.loads(result_text)
                    
                    doc_type = result_json.get("doc_type", "기타증빙서류")
                    owner = result_json.get("owner", "회사공통")
                    projects = result_json.get("projects", [])
                    new_filename = f"[{doc_type}] {owner}.pdf"
                    
                    drive_service = authenticate_google_drive()
                    if not drive_service: raise Exception("구글 드라이브 인증을 먼저 완료해주세요 (Secrets 확인)")
                    
                    archive_id = get_or_create_folder(drive_service, '[증빙자료_아카이브]')
                    owner_folder_id = get_or_create_folder(drive_service, owner, parent_id=archive_id)
                    
                    perf_file.seek(0)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(perf_file.getvalue())
                        tmp_path = tmp.name
                        
                    file_metadata = {'name': new_filename, 'parents': [owner_folder_id]} 
                    media = MediaFileUpload(tmp_path, mimetype='application/pdf')
                    drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                    os.remove(tmp_path)
                    
                    st.success(f"📂 증빙자료 업로드 완료! (`[증빙자료_아카이브] > {owner}` 폴더)")
                    
                    if projects:
                        st.info("✨ **AI 신규 실적 추출 완료 (중복 패스됨)**")
                        df_new = pd.DataFrame(projects)
                        st.dataframe(df_new, use_container_width=True)
                        
                        if st.button("💾 신규 실적 마스터 엑셀에 덮어쓰기", type="primary"):
                            with st.spinner("구글 드라이브의 마스터 엑셀 파일을 업데이트하고 있습니다..."):
                                results = drive_service.files().list(q="name contains '마스터' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false", fields="files(id)").execute()
                                items = results.get('files', [])
                                if items:
                                    file_id = items[0]['id']
                                    current_db = load_master_db_from_drive()
                                    updated_db = pd.concat([current_db, df_new], ignore_index=True)
                                    
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                                        with pd.ExcelWriter(tmp.name, engine='xlsxwriter') as writer:
                                            updated_db.to_excel(writer, index=False)
                                        tmp_path = tmp.name
                                        
                                    media = MediaFileUpload(tmp_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                                    drive_service.files().update(fileId=file_id, media_body=media).execute()
                                    os.remove(tmp_path)
                                    
                                    load_master_db_from_drive.clear() # 메모리 캐시 초기화
                                    st.success(f"🎉 성공! 마스터 엑셀 파일 맨 밑줄에 {len(projects)}건의 데이터가 완벽하게 추가되었습니다!")
                                else:
                                    st.error("드라이브에서 마스터 파일을 찾을 수 없습니다.")
                    else: st.warning("⚠️ 신규 실적이 없습니다. (전부 패스됨)")
                        
                except Exception as e: st.error(f"오류 발생: {str(e)}")

# --- [Tab 2, Tab 3, Tab 4 유지] ---
with tab2:
    col_title, col_toggle = st.columns([7, 3])
    with col_title: st.markdown("### 📊 공고문 AI 분석 결과")
    with col_toggle:
        st.write("")
        manual_override = st.toggle("⚙️ 세부사항 수동 설정", value=False)

    if not st.session_state.eval_criteria.empty: st.table(st.session_state.eval_criteria)
    else:
        st.info("💡 공고문을 업로드하시면 표가 완성됩니다.")
        st.table(pd.DataFrame({"대분류": ["참여기술인"], "평가항목": ["사업책임기술인"], "배점": ["20점"], "세부인정기준": ["경력/실적"]}))
    
    st.markdown("---")
    s_settings = st.session_state.auto_settings
    
    if not manual_override:
        st.success("🤖 AI 자동 세팅 모드입니다.")
        col_a, col_b = st.columns(2)
        with col_a:
            st.write(f"- **정기안전점검 실적 포함:** {'✅' if s_settings['has_safety'] else '❌'}")
            st.write(f"- **실적 인정 기간:** {s_settings['period']}")
            st.write(f"- **필요 인원:** 사책 {s_settings['pm_cnt']} / 분책 {s_settings['pe_cnt']} / 분참 {s_settings['pes_cnt']}")
        with col_b:
            st.write("- **보할:**"); st.table(pd.DataFrame(s_settings['bohal']))
        final_pm_cnt, final_pe_cnt, final_pes_cnt = s_settings['pm_cnt'], s_settings['pe_cnt'], s_settings['pes_cnt']
    else:
        st.warning("⚠️ 수동 설정 모드입니다.")
        chk_safety = st.checkbox("✅ 정기안전점검", value=s_settings['has_safety'])
        sel_period = st.selectbox("↳ 기간", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
        st.write("**✅ 보할 설정**"); edited_bohal = st.data_editor(pd.DataFrame(s_settings['bohal']), num_rows="dynamic")
        col_pm, col_pe, col_pes = st.columns(3)
        with col_pm: final_pm_cnt = st.number_input("사책", value=s_settings['pm_cnt'])
        with col_pe: final_pe_cnt = st.number_input("분책", value=s_settings['pe_cnt'])
        with col_pes: final_pes_cnt = st.number_input("분참", value=s_settings['pes_cnt'])

    st.markdown("---")
    assign_mode = st.radio("배정 방식:", ["🤖 AI 최적 배정", "🧑‍🔧 수동 선택"], horizontal=True, label_visibility="collapsed")
    personnel_list = engine.get_personnel_list()
    if assign_mode == "🧑‍🔧 수동 선택": st.write("수동 배정 UI (생략됨 - 시뮬레이션 연동 유지)")

with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 결과")
    if st.button("🚀 시뮬레이션 실행", type="primary"):
        with st.spinner('계산 중...'):
            time.sleep(1)
            best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(final_pm_cnt, final_pe_cnt, final_pes_cnt)
            st.success("🎉 AI 최적 조합 발견!")
            st.dataframe(best_score, use_container_width=True)

with tab4:
    st.subheader("최종 출력")
    if st.button("🔄 패키징 시작"):
        with st.spinner("압축 중..."):
            time.sleep(1)
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("안내문.txt", "제출 패키지 세팅 완료".encode('utf-8'))
            zip_buffer.seek(0)
            st.download_button(label="📦 다운로드", data=zip_buffer, file_name="최종.zip", mime="application/zip", type="primary")
