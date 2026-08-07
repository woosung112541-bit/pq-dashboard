import streamlit as st
import pandas as pd
import io
import json
import zipfile
import PyPDF2
import openpyxl # 💡 엑셀 원본 보존을 위한 필수 라이브러리
import google.generativeai as genai  
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# ⚙️ [초기 세팅]
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")

if 'auto_settings' not in st.session_state:
    st.session_state.auto_settings = {
        "has_safety": True, "period": "3년", "bohal": [{"전문분야": "상하수도", "비율(%)": 60}],
        "pm_cnt": 1, "pe_cnt": 2, "pes_cnt": 2, "extra_settings": {}
    }
if 'semi_fixed' not in st.session_state:
    st.session_state.semi_fixed = {
        "credit_rating": "BB-", "penalty_points": "해당없음", "new_hire_rate": 7.0,          
        "overlap_level": "최고 (350% 이상)", "investment_ratio": 3.03, "patent_tech_count": 5,        
        "bid_restriction": "해당없음", "inspection_penalty": "해당없음",  
    }
if 'semi_fixed_confirmed' not in st.session_state: st.session_state.semi_fixed_confirmed = False
if 'dream_team' not in st.session_state: st.session_state.dream_team = []
if 'notice_text' not in st.session_state: st.session_state.notice_text = ""
if 'raw_excel_bytes' not in st.session_state: st.session_state.raw_excel_bytes = b"" # 💡 엑셀 원본 파일 바이트 저장
if 'final_excel_bytes' not in st.session_state: st.session_state.final_excel_bytes = b"" # 완성된 엑셀 바이트

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
            token=None, refresh_token=oauth_data["refresh_token"], token_uri="https://oauth2.googleapis.com/token",
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
        request = drive_service.files().get_media(fileId=file_id) if items[0].get('mimeType', '') != 'application/vnd.google-apps.spreadsheet' else drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        while not downloader.next_chunk()[1]: pass
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
                elif file['mimeType'] == 'application/vnd.google-apps.folder': found_files.extend(get_all_files_recursively(drive_service, file['id'], target_mime_types, depth + 1))
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
# 🧠 [Backend Engine] 엑셀 좌표 다이렉트 주입 엔진
# ==========================================
class PQScoringEngine:
    def parse_excel_structure(self, excel_bytes):
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
        ws = wb.active
        structure = []
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            row_data = [str(cell) if cell is not None else "" for cell in row]
            if any(row_data): # 빈 줄 제외
                structure.append({"row_num": row_idx, "content": " | ".join(row_data)})
        return structure

    def run_ai_dreamteam_optimizer(self, notice_text, excel_bytes, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return None, [], [], []
            
        db_csv = master_db.to_csv(index=False)
        excel_structure = self.parse_excel_structure(excel_bytes)
        excel_json_str = json.dumps(excel_structure, ensure_ascii=False, indent=2)
        
        prompt = f"""
        당신은 건설엔지니어링 PQ '데이터 정밀 채점 매크로'입니다.
        제공된 [자기평가표 엑셀 구조]를 분석하고, [마스터 DB] 데이터를 기반으로 수학적 연산을 수행하여 각 'row_num(엑셀의 실제 행 번호)'에 삽입할 <획득점수>와 <산출근거>를 도출하세요.

        [★★★★★ 절대 준수 룰 - 데이터 매칭의 핵심]
        1. **지정공고문 우선:** [지정공고문]과 [세부평가기준] 내용 충돌 시 [지정공고문] 우선 적용.
        2. **기술인 역할 강제 매칭:** 엑셀 구조 내에 '사업책임'이라는 단어가 있는 행은 마스터 DB의 PM(예: 윤석순), '분야별책임'은 PE(예: 김진규), '분야별참여'는 PES(예: 황흥만)의 데이터를 **반드시 연결**해서 계산하세요. 대충 넘어가거나 "직접 입력 필요"라고 포기하지 마세요.
        3. **팩트 기반 수학 연산 (만점 환각 절대 금지):** 
           - 경력(년) = 총 참여일수 / 365
           - 실적 가중치 = 교량/터널 100%, 기타 토목 80% (건축 제외)
           반드시 "step_by_step_calc"에서 이 계산을 보여준 뒤 세부기준표 배점 구간을 찾으세요.
        4. **신용평가등급:** 당사는 '{semi_fixed.get('credit_rating')}'입니다. 세부기준표에서 이 등급의 점수를 찾아 감점하세요.
        5. **산출근거 간결화:** 'reason' 필드에는 수식이나 말을 빼고 최종 팩트 수치만 적으세요. (예: "특급", "5.33년", "22건", "1,288 백만원", "88 건", "해당없음", "BB-", "3.03%", "350% 이상", "7%")
        6. **소계/총계 연산:** 엑셀 구조 상의 '계', '소계', '총계' 행은 하위 항목의 점수를 모두 더해서 'score'를 적고, 'reason'은 "-" 로 하세요.

        [사용자 확인 반고정 항목]
        - 신용평가등급: {semi_fixed.get('credit_rating')}
        - 신규고용율: {semi_fixed.get('new_hire_rate')}%
        - 업무중첩도 구간: {semi_fixed.get('overlap_level')}
        - 기술개발비율: {semi_fixed.get('investment_ratio')}%

        [지정공고문 및 세부기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 구조 (row_num은 실제 엑셀의 행 번호임)]
        {excel_json_str}

        [마스터 DB]
        {db_csv}

        오직 순수 JSON으로만 반환하세요.
        {{
            "row_results": [
                {{"row_num": 2, "step_by_step_calc": "전체 합계 도출", "score": "91", "reason": "-"}},
                {{"row_num": 5, "step_by_step_calc": "김진규 등급 확인: 특급 (만점 2점)", "score": "2", "reason": "특급"}},
                {{"row_num": 8, "step_by_step_calc": "윤석순 참여일수 1946일 / 365 = 5.33년. 5년이상~5.5년미만 구간 6.4점", "score": "6.4", "reason": "5.33년"}}
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
            
            # 💡 [핵심] 엑셀 원본(Bytes)을 열고, AI가 지정한 행(row_num)에 값을 직접 덮어씁니다.
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
            ws = wb.active
            
            # '자기평가 점수' 및 '산출근거' 열의 인덱스(컬럼 번호) 찾기
            score_col_idx, reason_col_idx = -1, -1
            for col_idx, cell in enumerate(ws[1], start=1): # 첫 번째 행(헤더) 검사
                val = str(cell.value).replace(' ', '')
                if '자기평가' in val or '점수' in val: score_col_idx = col_idx
                elif '산출근거' in val or '산출' in val: reason_col_idx = col_idx
            
            # 헤더에서 못 찾으면 대략 마지막 2개 컬럼으로 추정
            if score_col_idx == -1: score_col_idx = ws.max_column - 1
            if reason_col_idx == -1: reason_col_idx = ws.max_column
            
            # 데이터 덮어쓰기
            for res in ai_results:
                row_idx = res.get("row_num")
                if row_idx:
                    # 빈 값이 아니면 덮어쓰기
                    if res.get("score") != "":
                        ws.cell(row=row_idx, column=score_col_idx, value=res.get("score"))
                    if res.get("reason") != "":
                        ws.cell(row=row_idx, column=reason_col_idx, value=res.get("reason"))
            
            # 수정된 엑셀을 다시 Bytes로 저장
            out_bytes = io.BytesIO()
            wb.save(out_bytes)
            
            # 화면 표시용 단순 DF 생성 (로깅용)
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
st.caption("※ 엑셀 원본 셀 다이렉트 주입 (Format 1000% 보존 엔진)")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

# --- [Tab 1] ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고/양식 업로드")
        notice_files = st.file_uploader("PDF 공고문 및 Excel 자기평가표 업로드", type=['pdf', 'xlsx'], accept_multiple_files=True)
        if notice_files and api_key and st.button("🧠 업로드 문서 AI 분석", type="primary"):
            with st.spinner("엑셀 원본 바이트를 메모리에 고정 중입니다..."):
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
                    st.success("✅ 문서 분석 완료! 엑셀 원본 파일이 손상 없이 시스템에 락업(Lock-up) 되었습니다.")
                else:
                    st.error("엑셀 파일이 업로드되지 않았거나 읽을 수 없습니다.")

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
    st.write(f"- 기업 신용평가등급: **{st.session_state.semi_fixed['credit_rating']}**")
    if st.session_state.raw_excel_bytes:
        st.write("- **자기평가서 엑셀 원본:** 시스템 메모리에 안전하게 보관 중 (직접 덮어쓰기 예정)")
    else:
        st.warning("Tab 1에서 엑셀 파일을 업로드해 주세요.")

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

    if st.button("🚀 자기평가서 점수 산출 (엑셀 파일 직접 수정)", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        if not st.session_state.raw_excel_bytes:
            st.error("엑셀 템플릿이 없습니다. Tab 1에서 엑셀을 업로드하세요.")
        else:
            with st.spinner('마스터 DB를 연산하고, 엑셀 원본 파일의 정확한 셀(Cell) 좌표에 값을 덮어쓰고 있습니다...'):
                final_excel, log_df, pm, pe, pes = engine.run_ai_dreamteam_optimizer(st.session_state.notice_text, st.session_state.raw_excel_bytes, st.session_state.semi_fixed)
                if final_excel:
                    st.success("🎉 산출 완료! 엑셀 원본 파일에 값이 완벽하게 주입되었습니다.")
                    st.session_state.final_excel_bytes = final_excel
                    st.session_state.dream_team = pm + pe + pes
                    st.info(f"👉 **최종 선발 명단:** {', '.join(st.session_state.dream_team)}")
                    
                    st.markdown("#### 🤖 AI 연산 결과 로깅 (참고용)")
                    st.caption("아래 표는 AI가 엑셀의 몇 번째 줄에 어떤 값을 넣었는지 보여주는 로그입니다. 실제 완벽한 표는 Tab 4에서 다운로드하는 엑셀 파일에 들어있습니다.")
                    st.dataframe(log_df, use_container_width=True)

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
                        # 💡 [핵심] 판다스로 만든 가짜 엑셀이 아니라, openpyxl로 수정한 진짜 원본 엑셀을 저장
                        z.writestr("0_완성된_자기평가표(원본보존).xlsx", st.session_state.final_excel_bytes)
                        
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
