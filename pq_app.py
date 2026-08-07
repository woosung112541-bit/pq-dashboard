import streamlit as st
import pandas as pd
import io
import json
import zipfile
import PyPDF2
import openpyxl
import google.generativeai as genai  
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# ⚙️ [초기 세팅]
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")

if 'semi_fixed' not in st.session_state:
    st.session_state.semi_fixed = {
        "credit_rating": "BB-",       
        "penalty_points": "해당없음",   
        "new_hire_rate": 7.0,          
        "overlap_level": "최고 (350% 이상)",        
        "investment_ratio": 3.03,       
        "patent_tech_count": 5,        
        "bid_restriction": "해당없음",  
        "inspection_penalty": "해당없음",  
    }
if 'semi_fixed_confirmed' not in st.session_state: st.session_state.semi_fixed_confirmed = False
if 'dream_team' not in st.session_state: st.session_state.dream_team = []
if 'notice_text' not in st.session_state: st.session_state.notice_text = ""
if 'raw_excel_bytes' not in st.session_state: st.session_state.raw_excel_bytes = b""
if 'final_excel_bytes' not in st.session_state: st.session_state.final_excel_bytes = b""

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
        results = drive_service.files().list(q="name contains '마스터' and trashed=false", fields="files(id, mimeType)").execute()
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
    def parse_excel_structure(self, excel_bytes):
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
        ws = wb.active
        structure = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            row_str = " | ".join([str(c).strip() for c in row if c is not None and str(c).strip() != ""])
            if row_str:
                structure.append({"row_num": row_idx, "row_content": row_str})
        return structure

    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt, notice_text, excel_bytes, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return None, pd.DataFrame(), [], [], []
            
        db_csv = master_db.to_csv(index=False)
        available_engineers = list(scan_drive_archive_cached().keys())
        engineers_str = ", ".join(available_engineers) if available_engineers else "명단 없음"
        
        excel_structure = self.parse_excel_structure(excel_bytes)
        excel_json_str = json.dumps(excel_structure, ensure_ascii=False, indent=2)
        
        prompt = f"""
        당신은 건설엔지니어링 PQ '정밀 채점 시스템'입니다.
        제공된 [지정공고문]과 [세부평가기준]을 적용하여 [마스터 DB] 데이터를 연산한 뒤, [자기평가표 엑셀 구조]의 빈칸(점수, 산출근거)을 정확하게 채우세요.

        [절대 준수 6대 규칙]
        1. 지정공고문 우선: 공고문과 세부기준이 다를 경우 반드시 공고문을 우선합니다.
        2. 엑셀 좌표 보존: 제공된 엑셀 구조의 'row_num'을 절대 삭제하거나 합치지 마세요. 소계, 총계 줄도 하위 항목을 더해서 반드시 출력해야 합니다.
        3. 수학 연산 강제 적용: [마스터 DB]의 데이터를 읽어 다음을 엄격히 계산하세요. (임의의 숫자를 지어내지 마세요.)
           - 교량/터널 분야는 가중치 1.0 (100%), 그 외 기타 토목 분야는 가중치 0.8 (80%) 적용
           - 환산 경력(년) = (Σ 교량/터널 참여일수 * 1.0 + Σ 기타 참여일수 * 0.8) / 365
           - 환산 실적(건수) = Σ 교량/터널 건수 * 1.0 + Σ 기타 건수 * 0.8
           - 환산 실적(금액) = Σ 교량/터널 금액 * 1.0 + Σ 기타 금액 * 0.8
        4. 반고정 팩트 데이터 적용:
           - 기업 신용평가등급: {semi_fixed.get('credit_rating')} -> 세부기준표 회사채 기준에서 해당 구간을 찾아 감점 (BB-는 2.8점 등)
           - 신규고용율: {semi_fixed.get('new_hire_rate')}% -> 가점 구간 적용
           - 업무중첩도: {semi_fixed.get('overlap_level')} -> 구간 점수 적용 (최고 350%이상은 최하점 부여)
           - 기술개발투자비율: {semi_fixed.get('investment_ratio')}% -> 배점 적용
        5. 역할(Role) 강제 매칭: 엑셀 행의 "사업책임"에는 선발된 PM 1명의 수치, "분야별책임"에는 PE의 수치, "분야별참여"에는 PES의 수치를 매칭하세요. "유사용역"은 회사의 전체 5년 합산 실적입니다.
        6. 극강의 간결성: 'reason(산출근거)' 필드에는 "팩트 숫자"만 적으세요. (예: "5.33년", "7건", "1,288 백만원", "특급", "BB-", "350% 이상", "7.0%")

        [사용 가능한 기술자 명단]
        [{engineers_str}] (이 중에서 최적의 PM {pm_cnt}명, PE {pe_cnt}명, PES {pes_cnt}명을 선발)

        [세부평가기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 구조 (row_num 좌표)]
        {excel_json_str}

        [마스터 DB]
        {db_csv}

        오직 순수 JSON으로만 반환:
        {{
            "step_by_step_log": "DB 데이터 추출 및 1.0/0.8 가중치 적용 연산 과정 요약",
            "row_results": [
                {{"row_num": 1, "score": "91.0", "reason": "-"}},
                {{"row_num": 8, "score": "6.4", "reason": "5.33년"}}
            ],
            "pm": ["윤석순"],
            "pe": ["김진규"],
            "pes": ["황흥만"]
        }}
        """
        try:
            model = genai.GenerativeModel('gemini-3.6-flash')
            response = model.generate_content(prompt)
            result_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
            parsed = json.loads(result_text)
            
            ai_results = parsed.get("row_results", [])
            
            # 엑셀 원본(Bytes) 수정
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
            ws = wb.active
            
            score_col_idx, reason_col_idx = -1, -1
            for r_idx in range(1, min(6, ws.max_row + 1)):
                for c_idx, cell in enumerate(ws[r_idx], start=1):
                    val = str(cell.value).replace(' ', '').replace('\n','')
                    if '자기평가' in val or '점수' in val: score_col_idx = c_idx
                    elif '산출근거' in val or '산출' in val: reason_col_idx = c_idx
            
            if score_col_idx == -1: score_col_idx = ws.max_column - 1
            if reason_col_idx == -1: reason_col_idx = ws.max_column
            
            for res in ai_results:
                row_idx = res.get("row_num")
                if row_idx:
                    score = str(res.get("score", "")).strip()
                    reason = str(res.get("reason", "")).strip()
                    
                    if score != "": 
                        try:
                            # 숫자형 변환 처리
                            ws.cell(row=row_idx, column=score_col_idx, value=float(score))
                        except ValueError:
                            ws.cell(row=row_idx, column=score_col_idx, value=score)
                    if reason != "" and reason != "점수": 
                        ws.cell(row=row_idx, column=reason_col_idx, value=reason)
            
            out_bytes = io.BytesIO()
            wb.save(out_bytes)
            display_df = pd.DataFrame(ai_results)
            
            return out_bytes.getvalue(), display_df, parsed.get("pm", []), parsed.get("pe", []), parsed.get("pes", [])
            
        except Exception as e:
            st.error(f"연산 오류: {e}")
            return None, pd.DataFrame(), [], [], []

engine = PQScoringEngine()

# ==========================================
# 🖥️ [Frontend]
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 엑셀 좌표 직접 주입 시스템")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. DB/서류 관리", "⚙️ 2. 공고문 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

with tab1:
    st.subheader("Zone A: 공고(PDF) & 자기평가표(Excel) 업로드")
    notice_files = st.file_uploader("", type=['pdf', 'xlsx'], accept_multiple_files=True)
    if notice_files and api_key and st.button("🧠 업로드 문서 시스템 적용", type="primary"):
        with st.spinner("엑셀 뼈대 구조를 고정하고 있습니다..."):
            notice_temp = ""
            excel_bytes_temp = b""
            for file in notice_files:
                if file.name.lower().endswith('.pdf'):
                    pdf = PyPDF2.PdfReader(file)
                    for page in pdf.pages[:20]: notice_temp += page.extract_text() or ""
                elif file.name.lower().endswith('.xlsx'):
                    excel_bytes_temp = file.getvalue()
            
            if excel_bytes_temp:
                st.session_state.notice_text = notice_temp
                st.session_state.raw_excel_bytes = excel_bytes_temp
                st.success("✅ 문서 분석 완료 및 엑셀 템플릿 적용 완료.")
            else: st.error("엑셀 파일이 업로드되지 않았습니다.")

with tab2:
    st.markdown("### 📊 설정 확인")
    if st.session_state.raw_excel_bytes:
        st.success("✅ 자기평가서 엑셀 템플릿 로드 완료.")
    else: st.warning("Tab 1에서 엑셀을 업로드하세요.")

with tab3:
    st.markdown("### 🏆 점수 산출")
    
    with st.expander("🔒 반고정 항목 적용 설정", expanded=not st.session_state.semi_fixed_confirmed):
        sf = st.session_state.semi_fixed
        col1, col2 = st.columns(2)
        with col1:
            sf['credit_rating'] = st.selectbox("신용평가등급 (예: BB-)", ["A-","BBB+","BBB0","BBB-","BB+","BB0","BB-","B+","B0","B-","CCC+"], index=6)
            sf['new_hire_rate'] = st.number_input("신규고용율 (%)", value=7.0, step=0.1)
        with col2:
            sf['overlap_level'] = st.selectbox("중첩도 (350% 이상)", ["최저", "중간", "최고 (350% 이상)"], index=2)
            sf['investment_ratio'] = st.number_input("투자비율 (%)", value=3.03, step=0.01)
        st.session_state.semi_fixed = sf
        st.session_state.semi_fixed_confirmed = st.checkbox("항목 확인 완료", value=st.session_state.semi_fixed_confirmed)

    if st.button("🚀 자기평가표 점수 자동 기입", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        if not st.session_state.raw_excel_bytes:
            st.error("엑셀 템플릿이 없습니다.")
        else:
            with st.spinner('마스터 DB를 연산하여 엑셀 원본 파일에 좌표를 덮어쓰고 있습니다...'):
                final_excel, log_df, pm, pe, pes = engine.run_ai_dreamteam_optimizer(1, 2, 2, st.session_state.notice_text, st.session_state.raw_excel_bytes, st.session_state.semi_fixed)
                if final_excel:
                    st.success("🎉 산출 완료! Tab 4에서 완성된 엑셀 파일을 다운로드하세요.")
                    st.session_state.final_excel_bytes = final_excel
                    st.session_state.dream_team = pm + pe + pes
                    st.dataframe(log_df, use_container_width=True)

with tab4:
    st.markdown("### 🖨️ 서류 출력 및 패키징")
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
                        z.writestr("0_완성된_자기평가표.xlsx", st.session_state.final_excel_bytes)
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
