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
# 🔑 [Google Drive 연동 및 데이터 전처리]
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

# 💡 [핵심] AI의 수학 계산 환각을 막기 위해 파이썬(Pandas)으로 미리 합계를 내어 프롬프트에 주입
def generate_db_summary_for_llm(master_df):
    if master_df.empty: return "마스터 DB 데이터 없음"
    summary = "=== [마스터 DB 자동 계산 요약 (수학 연산 환각 방지용 팩트 데이터)] ===\n"
    
    name_col = None
    for c in master_df.columns:
        if '성명' in str(c) or '이름' in str(c) or '기술인' in str(c):
            name_col = c; break
            
    if name_col:
        for name, group in master_df.groupby(name_col):
            summary += f"\n▶ 기술인: {name} (총 {len(group)}건의 프로젝트 수행)\n"
            num_df = group.select_dtypes(include=['number'])
            if not num_df.empty:
                sums = num_df.sum()
                for col in num_df.columns:
                    val = sums[col]
                    summary += f" - [{col}] 총합: {val}\n"
                    if '일' in col or '기간' in col:
                        summary += f"   (※ 365일 환산 경력: {val/365:.2f}년)\n"
                    if '금액' in col or '원' in col:
                        summary += f"   (※ 합계 금액: {val:,.0f})\n"
    else:
        summary += master_df.to_string(index=False)
    return summary

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
# 🧠 [Backend Engine] 엑셀 좌표 다이렉트 주입 엔진
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

    def run_ai_dreamteam_optimizer(self, notice_text, excel_bytes, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return None, pd.DataFrame(), [], [], []
            
        # 파이썬으로 계산 완료된 팩트 데이터 생성
        db_summary = generate_db_summary_for_llm(master_db)
        excel_structure = self.parse_excel_structure(excel_bytes)
        excel_json_str = json.dumps(excel_structure, ensure_ascii=False, indent=2)
        
        prompt = f"""
        당신은 건설엔지니어링 PQ '데이터 정밀 입력 봇'입니다.
        제공된 [자기평가표 엑셀 구조]를 분석하고, [마스터 DB 자동 계산 요약]의 데이터를 [세부평가기준]의 배점 구간에 대입하여, 각 'row_num(실제 엑셀 행 번호)'에 삽입할 <획득점수>와 <산출근거>를 도출하세요.

        [★★★★★ 절대 준수 규칙]
        1. **엑셀 원본 행(row_num) 완벽 보존:** 제공된 엑셀 구조의 row_num을 Key로 사용하세요. 행을 합치거나 지우면 시스템이 붕괴됩니다.
        2. **수학 계산 금지 (환각 방지):** 당신이 직접 숫자를 더하거나 나누지 마세요! 제공된 [마스터 DB 자동 계산 요약]에 있는 '환산 경력(년)', '건수 합계', '금액 합계' 등의 팩트 수치를 그대로 사용하여 세부기준 배점표 구간(예: 5년이상~5.5년미만 -> 6.4점)에만 매칭시키세요.
        3. **신용평가등급:** 당사는 '{semi_fixed.get('credit_rating')}'입니다. 세부기준표에서 해당 등급(예: B-이상~BBB-미만 등)의 점수를 찾아 감점하세요.
        4. **산출근거 극도 간결화:** 'reason' 필드에는 수식이나 말을 빼고 팩트만 적으세요. (예: "특급", "5.33년", "22건", "1,288 백만원", "88 건", "해당없음", "BB-", "3.03%", "350% 이상", "7%")
        5. **소계/총계 연산:** 엑셀 구조 상의 '계', '소계', '총계' 행은 도출된 하위 항목의 획득점수를 더해서 적고, 'reason'은 "-" 로 통일하세요.
        6. **데이터 없음 처리:** 데이터가 매칭되지 않는 빈 제목 행 등은 score와 reason을 ""(빈 문자열)로 두세요.

        [사용자 확인 반고정 항목 (이 값을 우선 적용)]
        - 신용평가등급: {semi_fixed.get('credit_rating')}
        - 신규고용율: {semi_fixed.get('new_hire_rate')}%
        - 업무중첩도 구간: {semi_fixed.get('overlap_level')}
        - 기술개발비율: {semi_fixed.get('investment_ratio')}%

        [세부평가기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 구조 (row_num 좌표)]
        {excel_json_str}

        [마스터 DB 자동 계산 요약 (수학 연산 오작동 방지용)]
        {db_summary}

        오직 순수 JSON으로만 반환:
        {{
            "row_results": [
                {{"row_num": 1, "score": "91.0", "reason": "-"}},
                {{"row_num": 8, "score": "6.4", "reason": "5.33년"}},
                {{"row_num": 9, "score": "", "reason": ""}} // 점수가 안 들어가는 제목 행
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
            
            # 💡 [핵심] 엑셀 원본(Bytes)을 열고, AI가 지정한 행(row_num)에 값을 직접 덮어쓰기
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
            ws = wb.active
            
            # '자기평가 점수' 및 '산출근거' 열을 자동으로 찾기
            score_col_idx, reason_col_idx = -1, -1
            for r_idx in range(1, min(5, ws.max_row + 1)):
                for c_idx, cell in enumerate(ws[r_idx], start=1):
                    val = str(cell.value).replace(' ', '').replace('\n','')
                    if '자기평가' in val or '점수' in val: score_col_idx = c_idx
                    elif '산출근거' in val or '산출' in val: reason_col_idx = c_idx
            
            # 실패 시 기본값 (맨 우측 2개 컬럼)
            if score_col_idx == -1: score_col_idx = ws.max_column - 1
            if reason_col_idx == -1: reason_col_idx = ws.max_column
            
            for res in ai_results:
                row_idx = res.get("row_num")
                if row_idx:
                    score = res.get("score", "")
                    reason = res.get("reason", "")
                    # 기존 헤더명 보존을 위해 숫자가 들어올 자리만 덮어씀
                    if score != "" and str(score).replace('.', '', 1).isdigit(): 
                        ws.cell(row=row_idx, column=score_col_idx, value=float(score) if '.' in str(score) else int(score))
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
st.caption("※ 엑셀 원본 셀 직접 주입 (100% 포맷 보존 엔진)")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고/양식 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

# --- [Tab 1] ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고/양식 업로드")
        notice_files = st.file_uploader("PDF 세부기준 및 Excel 자기평가표 업로드", type=['pdf', 'xlsx'], accept_multiple_files=True)
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
    if st.session_state.raw_excel_bytes:
        st.success("✅ 자기평가서 엑셀 원본이 메모리에 안전하게 보관 중입니다. (Tab 3에서 원본 파일에 팩트 데이터만 주입됩니다.)")
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
        st.session_state.semi_fixed_confirmed = st.checkbox("위 팩트 데이터(반고정 항목) 적용 확인", value=st.session_state.semi_fixed_confirmed)

    if st.button("🚀 자기평가표 점수 자동 기입 (엑셀 파일 직접 수정)", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        if not st.session_state.raw_excel_bytes:
            st.error("엑셀 템플릿이 없습니다. Tab 1에서 엑셀을 업로드하세요.")
        else:
            with st.spinner('마스터 DB를 연산하고, 엑셀 원본 파일의 정확한 셀(Cell) 좌표에 값을 덮어쓰고 있습니다...'):
                final_excel, log_df, pm, pe, pes = engine.run_ai_dreamteam_optimizer(st.session_state.notice_text, st.session_state.raw_excel_bytes, st.session_state.semi_fixed)
                if final_excel:
                    st.success("🎉 산출 완료! 엑셀 원본 파일에 값이 완벽하게 주입되었습니다. (Tab 4에서 다운로드하세요)")
                    st.session_state.final_excel_bytes = final_excel
                    st.session_state.dream_team = pm + pe + pes
                    st.info(f"👉 **최종 선발 명단:** {', '.join(st.session_state.dream_team)}")
                    
                    st.markdown("#### 🤖 AI 연산 결과 좌표 로깅")
                    st.caption("아래 표는 엑셀의 몇 번째 줄(row_num)에 어떤 값이 주입되었는지 확인하는 로그입니다. 엑셀 원본 렌더링 시 발생하는 깨짐 현상을 방지하기 위해 로깅 표만 화면에 제공합니다.")
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
                        # 💡 [핵심] openpyxl로 수정된 진짜 원본 엑셀 파일을 저장!
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
