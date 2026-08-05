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
# ⚙️ [초기 세팅] 시스템 가상 메모리 및 페이지 설정
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
if 'notice_text' not in st.session_state: st.session_state.notice_text = ""

with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")
    else:
        st.warning("문서 자동 분석을 위해 API Key가 필요합니다.")

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
        st.error(f"구글 드라이브 인증 실패: {e}")
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
        
        request = drive_service.files().get_media(fileId=items[0]['id'])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: _, done = downloader.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh)
    except Exception as e:
        return pd.DataFrame()

# 💡 [핵심 업그레이드 1] 특정 폴더 내부의 모든 PDF를 끝까지 파고들어(재귀탐색) 가져오는 함수
def get_all_pdfs_recursively(drive_service, folder_id):
    pdfs = []
    query = f"'{folder_id}' in parents and trashed=false"
    page_token = None
    while True:
        response = drive_service.files().list(q=query, fields='nextPageToken, files(id, name, mimeType)', pageToken=page_token).execute()
        for file in response.get('files', []):
            if file['mimeType'] == 'application/pdf':
                pdfs.append(file)
            elif file['mimeType'] == 'application/vnd.google-apps.folder':
                # 하위 폴더를 발견하면 그 안으로 한 번 더 진입!
                pdfs.extend(get_all_pdfs_recursively(drive_service, file['id']))
        page_token = response.get('nextPageToken', None)
        if not page_token: break
    return pdfs

# 💡 [핵심 업그레이드 2] 특정 폴더 내부의 모든 하위 폴더들의 이름과 ID를 맵핑하는 함수 (Tab 4에서 사람 찾기용)
def get_all_subfolders_map(drive_service, root_id):
    folder_dict = {}
    query = f"'{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    page_token = None
    while True:
        response = drive_service.files().list(q=query, fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
        for f in response.get('files', []):
            folder_dict[f['name']] = f['id']
            # 폴더 안의 폴더도 맵핑에 추가
            folder_dict.update(get_all_subfolders_map(drive_service, f['id']))
        page_token = response.get('nextPageToken', None)
        if not page_token: break
    return folder_dict

def scan_drive_archive():
    drive_service = authenticate_google_drive()
    if not drive_service: return {}
    q_arch = "name='[증빙자료_아카이브]' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res_arch = drive_service.files().list(q=q_arch, fields="files(id)").execute()
    if not res_arch.get('files'): return {}
    
    archive_id = res_arch.get('files')[0]['id']
    q_sub = f"'{archive_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    res_sub = drive_service.files().list(q=q_sub, fields="files(id, name)").execute()
    
    archive_status = {}
    for folder in res_sub.get('files', []):
        # 💡 [적용점] 최상위 폴더(예: 기술인)를 던져주면 그 안의 하위 트리를 모두 뒤져서 PDF만 싹 가져옴
        all_pdfs = get_all_pdfs_recursively(drive_service, folder['id'])
        archive_status[folder['name']] = [f['name'] for f in all_pdfs]
    return archive_status

def get_ai_model():
    return genai.GenerativeModel('gemini-3.6-flash')

# ==========================================
# 🧠 [Backend Engine] AI 다이렉트 시뮬레이션 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        self.master_db = load_master_db_from_drive()

    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt, notice_text):
        if self.master_db.empty:
            st.error("구글 드라이브에서 마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return pd.DataFrame(), [], [], []
            
        db_csv = self.master_db.to_csv(index=False)
        prompt = f"""
        당신은 건설엔지니어링 PQ(사업수행능력평가) 최고 심사위원입니다.
        아래 [공고문 세부기준]과 [엔지니어 실적 Master DB]를 분석하여, PQ 총점이 가장 높은 최적의 '드림팀'을 선발하세요.

        [평가 핵심 지침]
        1. 기간 필터링: 공고 기간이 3년일 경우, 현재(2026년) 기준으로 정확히 2025, 2024, 2023년도의 실적만 유효한 것으로 필터링하세요.
        2. 공고문 특화 기준 적용: '시안법', '건진법' 등 특정 법령 위주인지, 특정 분야(예: 항만) 가점이 있는지 공고문 텍스트에서 파악하여 DB 실적 점수에 완벽히 반영하세요.
        3. 업무중첩도: 공고문에 명시된 업무중첩도 감점 기준을 대입하여, 실적이 많아도 감점이 커지는 오판을 막고 총점이 가장 높은 조합을 찾으세요.
        4. 요구 인원: 사업책임기술자(PM) {pm_cnt}명, 분야별책임기술자(PE) {pe_cnt}명, 분야별참여기술자(PES) {pes_cnt}명 선발.

        [공고문 세부기준 텍스트]
        {notice_text[:3000]} 

        [엔지니어 실적 Master DB (CSV 형식)]
        {db_csv}

        위 데이터를 종합적으로 연산하여, 선발된 인원 조합과 점수 산출 내역을 아래 JSON 포맷으로만 반환하세요.
        {{
            "best_score_df": [
                {{"평가항목": "사업책임기술자", "배점": 30, "획득점수": 29.5, "비고": "윤석순 (항만 가점 반영, 중첩 감점 없음)"}},
                {{"평가항목": "분야별책임기술자", "배점": 40, "획득점수": 38.0, "비고": "시안법 실적 100% 인정"}}
            ],
            "pm": ["이름1"],
            "pe": ["이름2", "이름3"],
            "pes": ["이름4", "이름5"]
        }}
        """
        
        try:
            model = get_ai_model()
            response = model.generate_content(prompt)
            result_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
            parsed = json.loads(result_text)
            df = pd.DataFrame(parsed.get("best_score_df", []))
            return df, parsed.get("pm", []), parsed.get("pe", []), parsed.get("pes", [])
        except Exception as e:
            st.error(f"AI 연산 중 오류 발생: {e}")
            return pd.DataFrame(), [], [], []

engine = PQScoringEngine()

# ==========================================
# 🖥️ [Frontend] 메인 대시보드 UI
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 실무 완벽 대응: 딥-서치(Deep Search) 장착, 하위 폴더 전체 인식 모드")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 책임기술자 시뮬레이션", "🖨️ 4. 서류 출력 및 패키징"])

# --- [Tab 1] 마스터 DB 및 드라이브 스캔 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF) 업로드", type=['pdf'], accept_multiple_files=True, key="zone_a")
        if notice_files and api_key and st.button("🧠 공고문 AI 분석 및 평가기준 구성", type="primary"):
            with st.spinner("AI가 공고문을 정독 중입니다..."):
                try:
                    notice_text = ""
                    for file in notice_files:
                        pdf = PyPDF2.PdfReader(file)
                        for page in pdf.pages[:7]: notice_text += page.extract_text() or ""
                    st.session_state.notice_text = notice_text
                    prompt = f"건설엔지니어링 PQ 공고문 분석 후 JSON 반환.\n1. eval_criteria: 배점표\n2. settings: {{ has_safety, period, bohal, pm_cnt, pe_cnt, pes_cnt }}\n공고문: {notice_text}\n순수 JSON만 출력."
                    response = get_ai_model().generate_content(prompt)
                    parsed_json = json.loads(response.text.strip().removeprefix("```json").removesuffix("```").strip())
                    st.session_state.eval_criteria = pd.DataFrame(parsed_json.get("eval_criteria", []))
                    st.session_state.auto_settings = parsed_json.get("settings", st.session_state.auto_settings)
                    st.success("✅ 공고문 분석 성공! 평가 기준 및 텍스트가 시스템에 저장되었습니다.")
                except Exception as e:
                    st.error(f"공고문 분석 실패: {e}")
                        
    with col2:
        st.subheader("Zone B: 구글 드라이브 실적 아카이브 현황")
        st.info("💡 실적증명서 및 수료증 PDF는 구글 드라이브 `[증빙자료_아카이브]` 폴더에 자유롭게 업로드하시면 됩니다.")
        if st.button("🔍 구글 드라이브 아카이브 현황 새로고침"):
            load_master_db_from_drive.clear()
            st.rerun()

        with st.spinner("구글 드라이브 딥-스캔 중... (폴더 구조 파악)"):
            archive_data = scan_drive_archive()
            if archive_data:
                st.success(f"📂 최상위 카테고리 스캔 완료!")
                for name, pdfs in archive_data.items():
                    with st.expander(f"📁 **{name}** (총 {len(pdfs)}개 서류 보관 중)"):
                        if pdfs:
                            for pdf in pdfs: st.write(f"- 📄 `{pdf}`")
                        else:
                            st.caption("해당 카테고리 내부에는 어떠한 PDF도 존재하지 않습니다.")
            else:
                st.warning("구글 드라이브에 `[증빙자료_아카이브]` 폴더가 없거나 완전히 비어 있습니다.")

# --- [Tab 2] 공고문 세부사항 설정 ---
with tab2:
    st.markdown("### 📊 공고문 AI 분석 결과 및 세부 설정")
    col_title, col_toggle = st.columns([7, 3])
    with col_toggle: manual_override = st.toggle("⚙️ 세부사항 수동 설정", value=False)

    if not st.session_state.eval_criteria.empty: st.table(st.session_state.eval_criteria)
    else: st.info("💡 [Tab 1]에서 공고문을 업로드하시면 평가 배점표가 자동으로 구성됩니다.")
    
    st.markdown("---")
    s_settings = st.session_state.auto_settings
    if not manual_override:
        st.success("🤖 AI 자동 세팅 모드입니다.")
        col_a, col_b = st.columns(2)
        with col_a:
            st.write(f"- **정기안전점검 포함:** {'✅' if s_settings['has_safety'] else '❌'}")
            st.write(f"- **실적 인정 기간:** {s_settings['period']}")
            st.write(f"- **필요 인원:** 사책 {s_settings['pm_cnt']} / 분책 {s_settings['pe_cnt']} / 분참 {s_settings['pes_cnt']}")
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
    if st.button("🚀 마스터 DB 딥러닝 시뮬레이션 실행", type="primary"):
        if not st.session_state.notice_text:
            st.warning("⚠️ [Tab 1]에서 공고문을 먼저 업로드해야 해당 기준에 맞는 연산이 가능합니다!")
        else:
            with st.spinner('AI가 마스터 DB와 공고문 기준을 매칭하여 최적의 조합을 연산 중입니다...'):
                best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(
                    final_pm_cnt, final_pe_cnt, final_pes_cnt, st.session_state.notice_text
                )
                
                if not best_score.empty:
                    st.success("🎉 AI 최적 조합 산출 완료!")
                    st.dataframe(best_score, use_container_width=True)
                    st.session_state.dream_team = rec_pm + rec_pe + rec_pes
                    st.info(f"👉 **최종 선발 명단:** {', '.join(st.session_state.dream_team)}")

# --- [Tab 4] 서류 출력 및 패키징 ---
with tab4:
    st.markdown("### 🖨️ 서류 출력 및 자동 패키징")
    if not st.session_state.dream_team:
        st.warning("⚠️ 먼저 [Tab 3]에서 시뮬레이션을 실행하여 기술자 드림팀을 선발해 주세요.")
    else:
        st.success(f"✅ 현재 선발된 기술자 명단: **{', '.join(st.session_state.dream_team)}**")
        if st.button("🔄 구글 드라이브에서 서류 수집 및 ZIP 패키징 시작", type="primary"):
            with st.spinner("드라이브 전역을 뒤져 해당 인원의 모든 증빙 서류를 추출 중입니다..."):
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
                        
                        # 💡 [적용점] [증빙자료_아카이브] 내부의 모든 폴더 맵핑을 떠서 사람이름 폴더가 어디에 있든 찾아냄!
                        folder_map = get_all_subfolders_map(drive_service, archive_id)
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                            for person_name in st.session_state.dream_team:
                                if person_name in folder_map:
                                    person_folder_id = folder_map[person_name]
                                    
                                    # 해당 기술자 폴더 내부를 끝까지 파고들어(재귀) 모든 PDF 수집
                                    person_pdfs = get_all_pdfs_recursively(drive_service, person_folder_id)
                                    
                                    if person_pdfs:
                                        for pdf_file in person_pdfs:
                                            request = drive_service.files().get_media(fileId=pdf_file['id'])
                                            fh = io.BytesIO()
                                            downloader = MediaIoBaseDownload(fh, request)
                                            done = False
                                            while not done: _, done = downloader.next_chunk()
                                            fh.seek(0)
                                            z.writestr(f"{person_name}/{pdf_file['name']}", fh.read())
                                            found_files_count += 1
                                    else:
                                        z.writestr(f"{person_name}/안내_서류없음.txt", "폴더는 있으나 내부에 PDF 서류가 없습니다.".encode('utf-8'))
                                else:
                                    z.writestr(f"{person_name}/안내_폴더없음.txt", f"구글 드라이브에 '{person_name}' 폴더가 어디에도 없습니다.".encode('utf-8'))
                        
                        zip_buffer.seek(0)
                        if found_files_count > 0:
                            st.success(f"🎉 성공! 총 {found_files_count}개의 증빙 PDF를 전방위 탐색으로 찾아 압축했습니다.")
                            st.download_button("📦 최종 제출서류 패키지 다운로드 (ZIP)", data=zip_buffer, file_name="최종_PQ제출서류_패키지.zip", mime="application/zip", type="primary")
                        else: st.warning("선발된 인원에 해당하는 PDF 서류를 찾지 못했습니다.")
                except Exception as e: st.error(f"서류 패키징 중 오류 발생: {e}")
