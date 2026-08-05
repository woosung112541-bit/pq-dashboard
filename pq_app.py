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
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# ⚙️ [초기 세팅] 페이지 및 세션 상태 설정
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")

if 'eval_criteria' not in st.session_state: st.session_state.eval_criteria = pd.DataFrame() 
if 'auto_settings' not in st.session_state:
    st.session_state.auto_settings = {
        "has_safety": True, "period": "3년",
        "bohal": [{"전문분야": "상하수도", "비율(%)": 60}, {"전문분야": "토질지질", "비율(%)": 40}],
        "pm_cnt": 1, "pe_cnt": 2, "pes_cnt": 2
    }
if 'dream_team' not in st.session_state: st.session_state.dream_team = []

with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")
    else:
        st.warning("공고문 자동 분석을 위해 API Key가 필요합니다.")

# ==========================================
# 🔑 [Google Drive API 핵심 연동 함수]
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
        st.error(f"구글 드라이브 인증 실패 (Secrets 확인 필요): {e}")
        return None

@st.cache_data(ttl=300)
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

# 드라이브 아카이브 내 폴더 및 파일 현황 조회
def scan_drive_archive():
    drive_service = authenticate_google_drive()
    if not drive_service: return {}
    
    # 1. [증빙자료_아카이브] 폴더 탐색
    q_arch = "name='[증빙자료_아카이브]' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res_arch = drive_service.files().list(q=q_arch, fields="files(id)").execute()
    arch_files = res_arch.get('files', [])
    if not arch_files: return {}
    
    archive_id = arch_files[0]['id']
    
    # 2. 하위 기술자 폴더 탐색
    q_sub = f"'{archive_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res_sub = drive_service.files().list(q=q_sub, fields="files(id, name)").execute()
    sub_folders = res_sub.get('files', [])
    
    archive_status = {}
    for folder in sub_folders:
        p_name = folder['name']
        p_id = folder['id']
        # 폴더 내 PDF 파일 목록
        q_pdf = f"'{p_id}' in parents and mimeType='application/pdf' and trashed=false"
        res_pdf = drive_service.files().list(q=q_pdf, fields="files(name)").execute()
        pdf_list = [f['name'] for f in res_pdf.get('files', [])]
        archive_status[p_name] = pdf_list
        
    return archive_status

# ==========================================
# 🧠 [Backend Engine] PQ 점수 계산 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        self.master_db = load_master_db_from_drive()

    def get_personnel_list(self):
        if self.master_db.empty: return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]
        for col in ['이름', '성명', '엔지니어명', '기술자명', '기술인']:
            if col in self.master_db.columns: return ["(선택)"] + self.master_db[col].dropna().unique().tolist()
        return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]

    def run_ai_dreamteam_optimizer(self, pm, pe, pes):
        best_score_df = pd.DataFrame({
            "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
            "배점": [30, 20, 10, -5],
            "획득점수": [30.0, 20.0, 10.0, 0.0],
            "비고": ["AI 최적화", "중첩도 0건", "A+ 등급", "해당없음"]
        })
        sample_names = ["윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]
        return best_score_df, sample_names[0:pm], sample_names[pm:pm+pe], sample_names[pm+pe:pm+pe+pes]

engine = PQScoringEngine()

def get_ai_model():
    return genai.GenerativeModel('gemini-3.6-flash')

# ==========================================
# 🖥️ [Frontend] 메인 대시보드 UI
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 구글 드라이브 직접 연동 모드: 드라이브 내 서류 자동 스캔 및 스마트 패키징")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 책임기술자 시뮬레이션", "🖨️ 4. 서류 출력 및 패키징"])

# --- [Tab 1] 마스터 DB 및 드라이브 스캔 ---
with tab1:
    col1, col2 = st.columns(2)
    
    # Zone A: 공고문/지침서 분석
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF) 업로드", type=['pdf'], accept_multiple_files=True, key="zone_a")
        if notice_files and api_key and st.button("🧠 공고문 AI 분석 및 평가기준 구성", type="primary"):
            with st.spinner("AI가 공고문을 읽고 배점표 및 평가 설정을 추출 중입니다..."):
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
                    st.success("✅ 공고문 분석 성공! [Tab 2]에 모든 평가 기준이 세팅되었습니다.")
                except Exception as e:
                    st.error(f"공고문 분석 실패: {e}")
                        
    # Zone B: 구글 드라이브 아카이브 현황 조회
    with col2:
        st.subheader("Zone B: 구글 드라이브 실적 아카이브 현황")
        st.info("💡 실적증명서 및 수료증 PDF는 구글 드라이브 `[증빙자료_아카이브]` 폴더에 자유롭게 업로드하시면 됩니다.")
        
        if st.button("🔍 구글 드라이브 아카이브 현황 새로고침"):
            load_master_db_from_drive.clear() # 메모리 캐시 초기화
            st.rerun()

        with st.spinner("구글 드라이브 스캔 중..."):
            archive_data = scan_drive_archive()
            if archive_data:
                st.success(f"📂 총 {len(archive_data)}명의 기술자 폴더가 확인되었습니다.")
                for name, pdfs in archive_data.items():
                    with st.expander(f"👤 **{name}** ({len(pdfs)}개 서류 보관 중)"):
                        if pdfs:
                            for pdf in pdfs:
                                st.write(f"- 📄 `{pdf}`")
                        else:
                            st.caption("보관된 PDF 서류가 없습니다.")
            else:
                st.warning("구글 드라이브 최상단에 `[증빙자료_아카이브]` 폴더가 없거나 비어 있습니다.")

# --- [Tab 2] 공고문 세부사항 설정 ---
with tab2:
    st.markdown("### 📊 공고문 AI 분석 결과 및 세부 설정")
    col_title, col_toggle = st.columns([7, 3])
    with col_toggle:
        manual_override = st.toggle("⚙️ 세부사항 수동 설정", value=False)

    if not st.session_state.eval_criteria.empty:
        st.table(st.session_state.eval_criteria)
    else:
        st.info("💡 [Tab 1]에서 공고문을 업로드하시면 평가 배점표가 자동으로 구성됩니다.")
        st.table(pd.DataFrame({"대분류": ["참여기술인"], "평가항목": ["사업책임기술인"], "배점": ["20점"], "세부인정기준": ["경력/실적"]}))
    
    st.markdown("---")
    s_settings = st.session_state.auto_settings
    
    if not manual_override:
        st.success("🤖 AI 자동 세팅 모드입니다.")
        col_a, col_b = st.columns(2)
        with col_a:
            st.write(f"- **정기안전점검 실적 포함:** {'✅' if s_settings['has_safety'] else '❌'}")
            st.write(f"- **실적 인정 기간:** {s_settings['period']}")
            st.write(f"- **필요 인원:** 사책 {s_settings['pm_cnt']}명 / 분책 {s_settings['pe_cnt']}명 / 분참 {s_settings['pes_cnt']}명")
        with col_b:
            st.write("- **보할 인정 비율:**"); st.table(pd.DataFrame(s_settings['bohal']))
        final_pm_cnt, final_pe_cnt, final_pes_cnt = s_settings['pm_cnt'], s_settings['pe_cnt'], s_settings['pes_cnt']
    else:
        st.warning("⚠️ 수동 설정 모드입니다.")
        chk_safety = st.checkbox("✅ 정기안전점검 포함", value=s_settings['has_safety'])
        sel_period = st.selectbox("↳ 인정 기간", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
        st.write("**✅ 보할 설정**"); edited_bohal = st.data_editor(pd.DataFrame(s_settings['bohal']), num_rows="dynamic")
        col_pm, col_pe, col_pes = st.columns(3)
        with col_pm: final_pm_cnt = st.number_input("사책(명)", value=s_settings['pm_cnt'])
        with col_pe: final_pe_cnt = st.number_input("분책(명)", value=s_settings['pe_cnt'])
        with col_pes: final_pes_cnt = st.number_input("분참(명)", value=s_settings['pes_cnt'])

# --- [Tab 3] 책임기술자 시뮬레이션 결과 ---
with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 및 드림팀 선발")
    if st.button("🚀 시뮬레이션 실행 (최적 드림팀 추출)", type="primary"):
        with st.spinner('마스터 DB 기반 최적 기술자 조합 계산 중...'):
            time.sleep(1)
            best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(final_pm_cnt, final_pe_cnt, final_pes_cnt)
            
            st.success("🎉 AI 최적 조합 산출 완료!")
            st.dataframe(best_score, use_container_width=True)
            
            # 선발된 인원 명단을 세션 상태에 저장 (Tab 4 연동용)
            st.session_state.dream_team = rec_pm + rec_pe + rec_pes
            
            st.info(f"👉 **최종 선발 명단:** {', '.join(st.session_state.dream_team)}")
            st.caption("이 명단을 바탕으로 [Tab 4]에서 구글 드라이브의 해당 서류들을 한 번에 패키징합니다.")

# --- [Tab 4] 서류 출력 및 패키징 ---
with tab4:
    st.markdown("### 🖨️ 서류 출력 및 자동 패키징")
    
    if not st.session_state.dream_team:
        st.warning("⚠️ 먼저 [Tab 3]에서 시뮬레이션을 실행하여 기술자 드림팀을 선발해 주세요.")
    else:
        st.success(f"✅ 현재 선발된 기술자 명단: **{', '.join(st.session_state.dream_team)}**")
        st.write("버튼을 누르면 구글 드라이브 `[증빙자료_아카이브]`에서 해당 기술자들의 PDF 파일만 자동으로 찾아 압축(.zip) 파일로 만듭니다.")
        
        if st.button("🔄 구글 드라이브에서 서류 수집 및 ZIP 패키징 시작", type="primary"):
            with st.spinner("구글 드라이브에서 해당 인원의 PDF 증빙 서류를 수집 중입니다..."):
                try:
                    drive_service = authenticate_google_drive()
                    archive_query = "name='[증빙자료_아카이브]' and mimeType='application/vnd.google-apps.folder' and trashed=false"
                    archive_res = drive_service.files().list(q=archive_query, fields="files(id)").execute()
                    
                    if not archive_res.get('files'):
                        st.error("구글 드라이브에 `[증빙자료_아카이브]` 폴더가 존재하지 않습니다.")
                    else:
                        archive_id = archive_res['files'][0]['id']
                        zip_buffer = io.BytesIO()
                        found_files_count = 0
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                            for person_name in st.session_state.dream_team:
                                # 해당 인원의 폴더 탐색
                                person_q = f"name='{person_name}' and '{archive_id}' in parents and trashed=false"
                                person_res = drive_service.files().list(q=person_q, fields="files(id)").execute()
                                
                                if person_res.get('files'):
                                    person_folder_id = person_res['files'][0]['id']
                                    
                                    # 해당 폴더 안의 모든 PDF 파일 수집
                                    pdf_q = f"'{person_folder_id}' in parents and mimeType='application/pdf' and trashed=false"
                                    pdf_res = drive_service.files().list(q=pdf_q, fields="files(id, name)").execute()
                                    
                                    for pdf_file in pdf_res.get('files', []):
                                        request = drive_service.files().get_media(fileId=pdf_file['id'])
                                        fh = io.BytesIO()
                                        downloader = MediaIoBaseDownload(fh, request)
                                        done = False
                                        while not done: _, done = downloader.next_chunk()
                                        fh.seek(0)
                                        
                                        # ZIP 파일 내에 '기술자이름/파일명.pdf' 구조로 저장
                                        z.writestr(f"{person_name}/{pdf_file['name']}", fh.read())
                                        found_files_count += 1
                                else:
                                    # 드라이브에 폴더가 없는 경우 알림 파일 저장
                                    z.writestr(f"{person_name}/안내_서류없음.txt", f"구글 드라이브 [증빙자료_아카이브] 폴더 내에 '{person_name}' 폴더가 없습니다.".encode('utf-8'))
                        
                        zip_buffer.seek(0)
                        
                        if found_files_count > 0:
                            st.success(f"🎉 성공! 총 {found_files_count}개의 증빙 PDF 서류를 구글 드라이브에서 패키징했습니다.")
                            st.download_button(
                                label="📦 최종 제출서류 패키지 다운로드 (ZIP)",
                                data=zip_buffer,
                                file_name="최종_PQ제출서류_패키지.zip",
                                mime="application/zip",
                                type="primary"
                            )
                        else:
                            st.warning("선발된 인원에 해당하는 PDF 서류를 구글 드라이브에서 찾지 못했습니다.")
                except Exception as e:
                    st.error(f"서류 패키징 중 오류 발생: {e}")
