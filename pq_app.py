import streamlit as st
import pandas as pd
import io
import json
import os
import zipfile
import PyPDF2
import google.generativeai as genai  
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# ⚙️ [초기 세팅]
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")

if 'eval_criteria' not in st.session_state: st.session_state.eval_criteria = pd.DataFrame() 
if 'auto_settings' not in st.session_state:
    st.session_state.auto_settings = {
        "has_safety": True, "period": "3년",
        "bohal": [{"전문분야": "상하수도", "비율(%)": 60}],
        "pm_cnt": 1, "pe_cnt": 2, "pes_cnt": 2,
        "extra_settings": {}
    }
if 'semi_fixed' not in st.session_state:
    st.session_state.semi_fixed = {
        "credit_rating": "BB-",       
        "penalty_points": "해당없음",   
        "new_hire_rate": 0.0,          
        "overlap_level": "최고 (350% 이상)",        
        "investment_ratio": 0.0,       
        "patent_tech_count": 0,        
        "bid_restriction": "해당없음",  
        "inspection_penalty": "해당없음",  
    }
if 'semi_fixed_confirmed' not in st.session_state: st.session_state.semi_fixed_confirmed = False
if 'dream_team' not in st.session_state: st.session_state.dream_team = []
if 'notice_text' not in st.session_state: st.session_state.notice_text = ""
if 'self_eval_template' not in st.session_state: st.session_state.self_eval_template = ""
if 'final_pq_score_table' not in st.session_state: st.session_state.final_pq_score_table = pd.DataFrame()

with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")

# ==========================================
# 🔑 [Google Drive 연동]
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
    except Exception: return None

@st.cache_data(ttl=300)
def load_master_db_from_drive():
    try:
        drive_service = authenticate_google_drive()
        if not drive_service: return pd.DataFrame()
        results = drive_service.files().list(
            q="name contains '마스터' and trashed=false", fields="files(id, mimeType)"
        ).execute()
        items = results.get('files', [])
        if not items: return pd.DataFrame()
        
        file_id = items[0]['id']
        mime_type = items[0].get('mimeType', '')
        if mime_type == 'application/vnd.google-apps.spreadsheet':
            request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            request = drive_service.files().get_media(fileId=file_id)
            
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        return pd.read_excel(fh)
    except Exception: return pd.DataFrame()

def get_all_files_recursively(drive_service, folder_id, target_mime_types, depth=0):
    if depth > 5: return [] 
    found_files = []
    page_token = None
    while True:
        try:
            response = drive_service.files().list(q=f"'{folder_id}' in parents and trashed=false", fields='nextPageToken, files(id, name, mimeType)', pageToken=page_token).execute()
            for file in response.get('files', []):
                if file['mimeType'] in target_mime_types: found_files.append(file)
                elif file['mimeType'] == 'application/vnd.google-apps.folder':
                    found_files.extend(get_all_files_recursively(drive_service, file['id'], target_mime_types, depth + 1))
            page_token = response.get('nextPageToken', None)
            if not page_token: break
        except Exception: break
    return found_files

def get_all_subfolders_map(drive_service, root_id, depth=0):
    if depth > 5: return {}
    folder_dict = {}
    page_token = None
    while True:
        try:
            response = drive_service.files().list(q=f"'{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
            for f in response.get('files', []):
                folder_dict[f['name']] = f['id']
                folder_dict.update(get_all_subfolders_map(drive_service, f['id'], depth + 1))
            page_token = response.get('nextPageToken', None)
            if not page_token: break
        except Exception: break
    return folder_dict

@st.cache_data(ttl=300)
def scan_drive_archive_cached():
    try:
        drive_service = authenticate_google_drive()
        if not drive_service: return {}
        res_arch = drive_service.files().list(q="name='기술인' and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id)").execute()
        if not res_arch.get('files'): return {}
        
        archive_id = res_arch.get('files')[0]['id']
        res_sub = drive_service.files().list(q=f"'{archive_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields="files(id, name)").execute()
        
        archive_status = {}
        for folder in res_sub.get('files', []):
            all_pdfs = get_all_files_recursively(drive_service, folder['id'], ['application/pdf'])
            archive_status[folder['name']] = [f['name'] for f in all_pdfs]
        return archive_status
    except Exception: return {}

# ==========================================
# 🧠 [Backend Engine]
# ==========================================
class PQScoringEngine:
    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt, notice_text, self_eval_template, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return pd.DataFrame(), [], [], []
            
        db_csv = master_db.to_csv(index=False)
        
        prompt = f"""
        당신은 어떠한 요약이나 가공도 허용되지 않는 '단순 데이터 입력 매크로'입니다.
        아래 [자기평가표 엑셀 양식(CSV)]의 모든 행(Row)과 열(Column)을 100% 동일하게 복사하여 JSON으로 변환하세요.

        [★★★★★ 절대 준수 룰]
        1. **행(Row) 병합/생략 절대 금지:** 제공된 CSV 양식에 있는 줄 수 그대로 출력해야 합니다. 항목을 임의로 묶거나(예: '실적건수 및 금액' 등) 요약하면 즉시 시스템이 파괴됩니다. 빈 칸이 있는 줄(예: 소계, 계)도 무조건 그대로 출력하세요.
        2. **컬럼(Key) 동적 생성:** JSON의 Key는 제가 임의로 정하지 않습니다. 원본 CSV의 헤더(예: '평가항목', '배점', '자기평가 점수', '산출근거' 등)를 그대로 Key로 사용하세요.
        3. 당신의 유일한 임무는 100% 복사된 뼈대 안에서 비어있는 '점수'와 '산출근거' 열만 [마스터 DB]와 [세부기준]을 바탕으로 정확히 계산해서 채우는 것입니다.
        4. 신용평가등급은 {semi_fixed.get('credit_rating')} 입니다. 세부기준표 구간을 확인하여 정확한 감점 점수를 기입하세요.
        5. 데이터가 없으면 0점 처리하고 근거에 "직접 입력 필요"라고 명시하세요.

        [사용자 확인 반고정 항목]
        - 신용평가등급: {semi_fixed.get('credit_rating')}
        - 신규고용율: {semi_fixed.get('new_hire_rate')}%
        - 업무중첩도 구간: {semi_fixed.get('overlap_level')}
        - 기술개발비율: {semi_fixed.get('investment_ratio')}%

        [공고문 및 세부기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 양식 (이 구조를 100% 그대로 복사할 것)]
        {self_eval_template}

        [마스터 DB]
        {db_csv}

        오직 순수 JSON으로만 반환:
        {{
            "pq_score_table": [
                // 원본 CSV의 컬럼명을 그대로 Key로 사용. 줄(Row) 개수 동일 유지.
                // 예: {{"평가항목": "...", "배점": "...", "자기평가 점수": "...", "산출근거": "..."}}
            ],
            "pm": [".."], "pe": [".."], "pes": [".."]
        }}
        """
        try:
            model = genai.GenerativeModel('gemini-3.6-flash')
            response = model.generate_content(prompt)
            result_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
            parsed = json.loads(result_text)
            return pd.DataFrame(parsed.get("pq_score_table", [])), parsed.get("pm", []), parsed.get("pe", []), parsed.get("pes", [])
        except Exception as e:
            st.error(f"연산 오류: {e}")
            return pd.DataFrame(), [], [], []

engine = PQScoringEngine()

# ==========================================
# 🖥️ [Frontend]
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 엑셀 원본 행(Row) 절대 보존 매크로 엔진 탑재")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

# --- [Tab 1] ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고/양식 업로드")
        notice_files = st.file_uploader("PDF 공고문 및 Excel 자기평가표 업로드", type=['pdf', 'xlsx'], accept_multiple_files=True)
        if notice_files and api_key and st.button("🧠 업로드 문서 AI 분석", type="primary"):
            with st.spinner("엑셀 양식을 100% 매핑하여 복제 중입니다..."):
                notice_temp, excel_temp = "", ""
                for file in notice_files:
                    if file.name.lower().endswith('.pdf'):
                        pdf = PyPDF2.PdfReader(file)
                        for page in pdf.pages[:18]: notice_temp += page.extract_text() or ""
                    elif file.name.lower().endswith('.xlsx'):
                        df_ex = pd.read_excel(file)
                        excel_temp += df_ex.to_csv(index=False)
                st.session_state.notice_text = notice_temp
                st.session_state.self_eval_template = excel_temp
                st.success("✅ 문서 매칭 완료! 엑셀 서식이 시스템에 강제 주입되었습니다.")

    with col2:
        st.subheader("Zone B: '기술인' 서류 스캔")
        if st.button("🔍 드라이브 재스캔"):
            scan_drive_archive_cached.clear()
            st.rerun()
        archive_data = scan_drive_archive_cached()
        if archive_data:
            st.success(f"📂 기술자 {len(archive_data)}명 스캔 완료!")

# --- [Tab 2] ---
with tab2:
    st.markdown("### 📊 세부 설정 확인")
    col_a, col_b = st.columns(2)
    with col_a:
        st.write(f"- 기업 신용평가등급: **{st.session_state.semi_fixed['credit_rating']}**")
        st.write("- 자기평가서 엑셀 양식: **강제 적용 완료** (요약 불가 상태)")

# --- [Tab 3] ---
with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 및 점수 산출")
    
    with st.expander("🔒 반고정 항목 확인 (필수)", expanded=not st.session_state.semi_fixed_confirmed):
        sf = st.session_state.semi_fixed
        col1, col2 = st.columns(2)
        with col1:
            sf['credit_rating'] = st.selectbox("신용평가등급", ["A-","BBB+","BBB0","BBB-","BB+","BB0","BB-","B+","B0","B-","CCC+"], index=6)
            sf['new_hire_rate'] = st.number_input("신규고용율 (%)", value=7.0, step=0.1)
        with col2:
            sf['overlap_level'] = st.selectbox("중첩도 수준", ["최저", "중간", "최고 (350% 이상)"], index=2)
            sf['investment_ratio'] = st.number_input("투자비율 (%)", value=3.03, step=0.01)
        st.session_state.semi_fixed = sf
        st.session_state.semi_fixed_confirmed = st.checkbox("항목 확인 완료", value=st.session_state.semi_fixed_confirmed)

    if st.button("🚀 자기평가서 엑셀 미러링 기반 점수 산출", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        with st.spinner('원본 엑셀의 줄(Row)을 그대로 복사하여 팩트 데이터를 채우고 있습니다...'):
            df_res, pm, pe, pes = engine.run_ai_dreamteam_optimizer(1, 2, 2, st.session_state.notice_text, st.session_state.self_eval_template, st.session_state.semi_fixed)
            if not df_res.empty:
                st.success("🎉 산출 완료!")
                st.dataframe(df_res.astype(str), use_container_width=True)
                st.session_state.final_pq_score_table = df_res
                st.session_state.dream_team = pm + pe + pes

# --- [Tab 4] ---
with tab4:
    st.markdown("### 🖨️ 자동 패키징")
    if st.session_state.dream_team:
        if st.button("🔄 증빙 서류 및 엑셀 결과 다운로드 (ZIP)", type="primary"):
            with st.spinner("압축 중..."):
                try:
                    drive_service = authenticate_google_drive()
                    arch_res = drive_service.files().list(q="name='기술인' and trashed=false", fields="files(id)").execute()
                    arch_id = arch_res['files'][0]['id']
                    f_map = get_all_subfolders_map(drive_service, arch_id)
                    
                    z_buf = io.BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as z:
                        e_buf = io.BytesIO()
                        st.session_state.final_pq_score_table.to_excel(e_buf, index=False)
                        z.writestr("0_완성된_자기평가표.xlsx", e_buf.getvalue())
                        
                        for p_name in st.session_state.dream_team:
                            if p_name in f_map:
                                pdfs = get_all_files_recursively(drive_service, f_map[p_name], ['application/pdf'])
                                for p_file in pdfs:
                                    req = drive_service.files().get_media(fileId=p_file['id'])
                                    fh = io.BytesIO()
                                    dl = MediaIoBaseDownload(fh, req)
                                    while not dl.next_chunk()[1]: pass
                                    fh.seek(0)
                                    z.writestr(f"{p_name}/{p_file['name']}", fh.read())
                    st.download_button("📦 다운로드", data=z_buf.getvalue(), file_name="최종패키지.zip", mime="application/zip")
                except Exception as e: st.error(e)
