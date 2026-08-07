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

def generate_db_summary_for_llm(master_df):
    if master_df.empty: return "마스터 DB 데이터 없음"
    summary = "=== [마스터 DB 자동 계산 요약 (수학 연산 환각 방지용 팩트 데이터)] ===\n"
    summary += "※ AI는 아래의 수치를 절대 임의로 변경하거나 재계산하지 말고, 오직 [세부평가기준]의 어느 배점 구간에 해당하는지만 찾아서 점수를 매길 것.\n\n"
    
    name_col = None
    for c in master_df.columns:
        if '성명' in str(c) or '이름' in str(c) or '기술인' in str(c):
            name_col = c; break
            
    if name_col:
        for name, group in master_df.groupby(name_col):
            summary += f"▶ 기술인: {name}\n"
            total_days = 0
            total_count = 0
            total_amount = 0
            for _, row in group.iterrows():
                weight = 1.0 if '교량' in str(row.values) or '터널' in str(row.values) else 0.8
                days = 0
                for v in row.values:
                    if isinstance(v, (int, float)) and v > 1000 and v < 10000:
                        days = v; break
                total_days += (days * weight)
                
                amt = 0
                for v in row.values:
                    if isinstance(v, (int, float)) and v > 10000:
                        amt = v; break
                total_amount += (amt * weight)
                total_count += (1 * weight)
                
            years = total_days / 365.0
            
            summary += f" - 환산 경력(년): {years:.2f}년\n"
            summary += f" - 환산 실적 건수: {total_count:.1f}건\n"
            summary += f" - 환산 실적 금액: {total_amount:,.0f} 원\n\n"
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
# 🧠 [Backend Engine] 팩트 주입형 매크로 엔진
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

    # 💡 [핵심] 병합된 셀에 안전하게 값을 쓰는 스마트 함수
    def write_cell_safe(self, sheet, r, c, val):
        cell = sheet.cell(row=r, column=c)
        if type(cell).__name__ == 'MergedCell':
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value = val
                    break
        else:
            cell.value = val

    def run_ai_dreamteam_optimizer(self, notice_text, excel_bytes, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return None, pd.DataFrame(), [], [], []
            
        fact_data_summary = generate_db_summary_for_llm(master_db)
        excel_structure = self.parse_excel_structure(excel_bytes)
        excel_json_str = json.dumps(excel_structure, ensure_ascii=False, indent=2)
        
        prompt = f"""
        당신은 수학 계산을 할 필요가 없는 '구간 매칭 매크로'입니다.
        아래 [파이썬 사전 계산 팩트 데이터]에 모든 수학적 연산 결과(경력 년수, 건수, 금액)가 나와 있습니다.
        당신의 임무는 오직 해당 팩트 수치가 [세부평가기준] 배점표의 어느 '구간'에 해당하는지 찾아내서, [자기평가표 엑셀 구조]의 알맞은 행(row_num)에 <점수>와 <팩트>를 적어넣는 것뿐입니다.

        [★★★★★ 절대 준수 규칙]
        1. **계산 금지:** 스스로 더하거나 나누지 마세요. [파이썬 사전 계산 팩트 데이터]의 숫자를 100% 믿고 그 숫자 그대로 세부기준표에서 점수만 찾으세요.
        2. **신용평가등급 절대 규칙:** 당사는 '{semi_fixed.get('credit_rating')}'입니다. 회사채 기준 세부기준표를 보면 BB- 등급은 무조건 **2.8점**입니다. (다른 점수 절대 불가)
        3. **업무중첩도 절대 규칙:** 구간이 '{semi_fixed.get('overlap_level')}'이므로, 세부기준 배점표에서 350% 이상 구간에 해당하는 **최하점(사업책임 3.6점, 분야별책임 2.4점)**을 무조건 적용하세요.
        4. **가점 절대 규칙:** 신규고용율이 {semi_fixed.get('new_hire_rate')}% 이므로 세부기준 가점표에서 최고구간인 **0.3점**을 무조건 적용하세요.
        5. **reason(산출근거) 작성법:** 수식을 빼고 최종 팩트 숫자만 아주 짧게 기입하세요. (예: "5.33년", "22건", "1,288 백만원", "특급", "BB-", "350% 이상", "7.0%")
        6. 엑셀 원본 행 번호(row_num) 유지, 소계/총계는 하위 항목 합산 후 reason은 "-" 기입.

        [세부평가기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 구조 (이 번호에 맞춰 점수 기입)]
        {excel_json_str}

        [파이썬 사전 계산 팩트 데이터 (이 수치들을 세부기준표 구간에 매칭할 것)]
        {fact_data_summary}

        오직 순수 JSON으로만 반환:
        {{
            "row_results": [
                {{"row_num": 8, "score": "6.4", "reason": "5.33년"}},
                {{"row_num": 9, "score": "3.6", "reason": "0.59년"}}
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
                    score = res.get("score", "")
                    reason = res.get("reason", "")
                    
                    if score != "" and str(score).replace('.', '', 1).isdigit(): 
                        self.write_cell_safe(ws, row_idx, score_col_idx, float(score) if '.' in str(score) else int(score))
                    if reason != "" and reason != "점수": 
                        self.write_cell_safe(ws, row_idx, reason_col_idx, reason)
            
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
st.caption("※ 파이썬 사전 계산 + 엑셀 직접 주입 엔진 (MergedCell 버그 해결판)")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. DB/서류 관리", "⚙️ 2. 공고문 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

with tab1:
    st.subheader("Zone A: 공고(PDF) & 자기평가표(Excel) 업로드")
    notice_files = st.file_uploader("", type=['pdf', 'xlsx'], accept_multiple_files=True)
    if notice_files and api_key and st.button("🧠 업로드 문서 AI 분석", type="primary"):
        with st.spinner("엑셀 원본 보존 및 기준 파악 중..."):
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
                st.success("✅ 문서 분석 완료! 엑셀 뼈대 고정됨.")
            else: st.error("엑셀 파일 없음.")

with tab2:
    st.markdown("### 📊 세부 설정 확인")
    if st.session_state.raw_excel_bytes:
        st.success("✅ 엑셀 원본 파일이 시스템 메모리에 로드되어 있습니다.")
    else: st.warning("Tab 1에서 엑셀을 업로드하세요.")

with tab3:
    st.markdown("### 🏆 점수 산출")
    
    with st.expander("🔒 팩트 데이터(반고정 항목) 적용 설정", expanded=not st.session_state.semi_fixed_confirmed):
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

    if st.button("🚀 점수 산출 (파이썬 사전 계산 방식)", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        if not st.session_state.raw_excel_bytes:
            st.error("엑셀 템플릿이 없습니다.")
        else:
            with st.spinner('시스템이 수학 계산을 완료하고 엑셀 원본 파일에 좌표를 찍어 덮어쓰고 있습니다...'):
                final_excel, log_df, pm, pe, pes = engine.run_ai_dreamteam_optimizer(st.session_state.notice_text, st.session_state.raw_excel_bytes, st.session_state.semi_fixed)
                if final_excel:
                    st.success("🎉 산출 완료! Tab 4에서 원본 엑셀 파일을 다운로드하세요.")
                    st.session_state.final_excel_bytes = final_excel
                    st.session_state.dream_team = pm + pe + pes
                    st.dataframe(log_df, use_container_width=True)

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
