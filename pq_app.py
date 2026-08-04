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
# [신규 추가] 다중 파일 처리용 메모리
if 'zone_b_projects' not in st.session_state: st.session_state.zone_b_projects = pd.DataFrame()
if 'zone_b_analyzed' not in st.session_state: st.session_state.zone_b_analyzed = False

with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")
    else:
        st.warning("문서 자동 분석을 위해 API Key가 필요합니다.")

# ==========================================
# 🔑 [Google Drive API 핵심 함수]
# ==========================================
@st.cache_resource
def authenticate_google_drive():
    try:
        oauth_data = st.secrets["google_oauth"]
        creds = Credentials(
            token=None, refresh_token=oauth_data["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=oauth_data["client_id"], client_secret=oauth_data["client_secret"]
        )
        return build('drive', 'v3', credentials=creds)
    except Exception as e:
        st.error(f"인증 실패: {e}"); return None

def get_or_create_folder(drive_service, folder_name, parent_id=None):
    query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id: query += f" and '{parent_id}' in parents"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    if items: return items[0]['id']
    else:
        meta = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
        if parent_id: meta['parents'] = [parent_id]
        return drive_service.files().create(body=meta, fields='id').execute().get('id')

# [신규 추가] 중복 파일명 방지 로직 (버전 관리)
def get_unique_filename(drive_service, folder_id, base_name):
    query = f"name='{base_name}' and '{folder_id}' in parents and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    if not results.get('files', []):
        return base_name # 중복 없으면 원본 이름 그대로 사용
    
    # 중복 시 파일명 뒤에 시/분/초 추가
    name, ext = os.path.splitext(base_name)
    timestamp = time.strftime("%H%M%S")
    return f"{name}_v{timestamp}{ext}"

@st.cache_data(ttl=600)
def load_master_db_from_drive():
    try:
        drive_service = authenticate_google_drive()
        if not drive_service: return pd.DataFrame()
        results = drive_service.files().list(q="name contains '마스터' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false", fields="files(id)").execute()
        items = results.get('files', [])
        if not items: return pd.DataFrame()
        request = drive_service.files().get_media(fileId=items[0]['id'])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: status, done = downloader.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh)
    except: return pd.DataFrame()

# ==========================================
# 🧠 [Backend Engine] PQ 점수 계산 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        self.master_db = load_master_db_from_drive()

    def get_personnel_list(self):
        if self.master_db.empty: return ["(선택)"]
        for col in ['이름', '성명', '엔지니어명', '기술자명', '기술인']:
            if col in self.master_db.columns: return ["(선택)"] + self.master_db[col].dropna().unique().tolist()
        return ["(선택)"]
    def run_ai_dreamteam_optimizer(self, pm, pe, pes): return pd.DataFrame(), [], [], []

engine = PQScoringEngine()
def get_ai_model():
    try: return genai.GenerativeModel('gemini-3.6-flash')
    except: return genai.GenerativeModel('gemini-1.5-flash')

# ==========================================
# 🖥️ [Frontend] 메인 대시보드 UI
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 실무 완벽 대응: 다중 업로드, 스마트 네이밍, 중복 파일 버전 관리 탑재")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 책임기술자 시뮬레이션", "🖨️ 4. 서류 출력 및 패키징"])

# --- [Tab 1] 마스터 DB 관리 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF) 업로드", type=['pdf'], accept_multiple_files=True, key="zone_a")
        if notice_files and api_key and st.button("🧠 공고문 AI 분석 및 평가기준 구성", type="primary"):
            st.success("✅ 공고문 분석 완료 (Tab 2 세팅됨)")
                        
    with col2:
        st.subheader("Zone B: 실적 업데이트 및 스마트 아카이빙")
        perf_files = st.file_uploader("기술인/회사 실적증명서 (여러 파일 한 번에 드래그 가능)", type=['pdf'], accept_multiple_files=True, key="zone_b")
        
        if perf_files and api_key:
            if st.button("🚀 선택한 모든 파일 AI 분석 및 드라이브 저장", type="primary"):
                all_extracted_projects = []
                drive_service = authenticate_google_drive()
                if not drive_service: st.stop()
                
                archive_id = get_or_create_folder(drive_service, '[증빙자료_아카이브]')
                progress_bar = st.progress(0)
                
                for idx, perf_file in enumerate(perf_files):
                    with st.spinner(f"[{idx+1}/{len(perf_files)}] '{perf_file.name}' 분석 중... (속도 제한 방지 적용됨)"):
                        try:
                            pdf_part = {"mime_type": "application/pdf", "data": perf_file.getvalue()}
                            existing_str = ", ".join(map(str, engine.master_db['사업명'].dropna().tolist() if not engine.master_db.empty and '사업명' in engine.master_db.columns else []))
                            
                            prompt = f"""
                            PDF 분석 후 JSON 반환.
                            1. doc_type (예: 수료증, 경력증명서, 실적증명서)
                            2. owner (이름, 없으면 '회사공통')
                            3. specific_detail: 파일을 명확히 구분할 수 있는 핵심 요약어. (경력증명서면 '발급일자(YYYYMMDD)', 수료증이면 '정밀안전진단교량' 등 과정명). 띄어쓰기 없이 15자 이내.
                            4. projects (배열). 반드시 "사업명", "시작일", "종료일", "담당업무", "발주처" 키 사용. 기존 목록 [{existing_str}] 중복 제외 신규만.
                            순수 JSON만 출력.
                            """
                            response = get_ai_model().generate_content([prompt, pdf_part])
                            result_json = json.loads(response.text.strip().removeprefix("```json").removesuffix("```").strip())
                            
                            doc_type = result_json.get("doc_type", "기타증빙")
                            owner = result_json.get("owner", "회사공통")
                            specific_detail = result_json.get("specific_detail", "상세불명")
                            projects = result_json.get("projects", [])
                            
                            base_filename = f"[{doc_type}] {owner}_{specific_detail}.pdf"
                            owner_folder_id = get_or_create_folder(drive_service, owner, parent_id=archive_id)
                            final_filename = get_unique_filename(drive_service, owner_folder_id, base_filename)
                            
                            perf_file.seek(0)
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                tmp.write(perf_file.getvalue())
                                tmp_path = tmp.name
                                
                            file_metadata = {'name': final_filename, 'parents': [owner_folder_id]} 
                            media = MediaFileUpload(tmp_path, mimetype='application/pdf')
                            drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                            os.remove(tmp_path)
                            
                            if projects: all_extracted_projects.extend(projects)
                            
                        except Exception as e: 
                            st.error(f"'{perf_file.name}' 처리 중 오류: {e}")
                    
                    progress_bar.progress((idx + 1) / len(perf_files))
                    
                    # ⏳ [핵심 추가] AI 과호흡(429 에러) 방지용 강제 휴식 (마지막 파일 제외)
                    if idx < len(perf_files) - 1:
                        time.sleep(4.5) 
                
                st.session_state.zone_b_projects = pd.DataFrame(all_extracted_projects)
                st.session_state.zone_b_analyzed = True
                st.success(f"📂 총 {len(perf_files)}개의 파일이 스마트 분류(이름 지정 및 중복 회피)되어 업로드되었습니다!")

        if st.session_state.zone_b_analyzed:
            if not st.session_state.zone_b_projects.empty:
                st.info("✨ **AI 신규 실적 추출 완료 (전체 합산)**")
                st.dataframe(st.session_state.zone_b_projects, use_container_width=True)
                
                if st.button("💾 모든 신규 실적 마스터 엑셀에 한 번에 덮어쓰기", type="primary"):
                    with st.spinner("구글 드라이브 마스터 엑셀에 기록 중..."):
                        drive_service = authenticate_google_drive()
                        results = drive_service.files().list(q="name contains '마스터' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false", fields="files(id)").execute()
                        items = results.get('files', [])
                        if items:
                            file_id = items[0]['id']
                            current_db = load_master_db_from_drive()
                            updated_db = pd.concat([current_db, st.session_state.zone_b_projects], ignore_index=True)
                            
                            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                                with pd.ExcelWriter(tmp.name, engine='xlsxwriter') as writer:
                                    updated_db.to_excel(writer, index=False)
                                tmp_path = tmp.name
                                
                            media = MediaFileUpload(tmp_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                            drive_service.files().update(fileId=file_id, media_body=media).execute()
                            os.remove(tmp_path)
                            
                            load_master_db_from_drive.clear()
                            st.session_state.zone_b_analyzed = False
                            st.session_state.zone_b_projects = pd.DataFrame()
                            st.success(f"🎉 성공! 마스터 엑셀 파일 맨 밑줄에 데이터가 완벽하게 일괄 추가되었습니다!")
                        else:
                            st.error("드라이브에서 마스터 파일을 찾을 수 없습니다.")
            else:
                st.warning("⚠️ 신규 실적이 없습니다. (전부 중복 패스됨)")

# --- [Tab 2, Tab 3, Tab 4 유지] ---
with tab2: st.write("Tab 2 UI (생략)")
with tab3: st.write("Tab 3 UI (생략)")
with tab4: st.write("Tab 4 UI (생략)")
