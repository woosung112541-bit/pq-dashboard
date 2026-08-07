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
if 'db_summary_cache' not in st.session_state: st.session_state.db_summary_cache = ""

with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")

# ==========================================
# 🔑 [Google Drive 연동 및 파이썬 자동 연산]
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
        if not drive_service: return {}
        results = drive_service.files().list(q="name contains '마스터' and trashed=false", fields="files(id, mimeType)").execute()
        items = results.get('files', [])
        if not items: return {}
        
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
        
        # 💡 [핵심수정] 모든 시트를 딕셔너리로 읽어옴 (시트명 = 사람이름)
        return pd.read_excel(fh, sheet_name=None)
    except Exception: return {}

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

# 💡 [핵심엔진 1] 파이썬을 이용한 스마트 팩트 데이터 추출기 (시트=이름, 행=건수)
def calculate_fact_data(master_db_dict):
    if not master_db_dict: return "마스터 DB 데이터 없음"
    summary = "=== [파이썬 시스템 자동 산출 팩트 데이터] ===\n"
    summary += "※ AI는 아래의 수치를 절대 임의로 변경하거나 재계산하지 말고, 오직 이 팩트 그대로 세부기준표에서 점수만 매칭할 것.\n\n"
    
    for sheet_name, df in master_db_dict.items():
        if df.empty or sheet_name.startswith('Sheet'): continue
        
        name = str(sheet_name).strip()
        summary += f"▶ 기술인: {name}\n"
        summary += f" - 등급: 특급 (고정)\n" # 💡 등급은 특급으로 고정
        
        # 컬럼 자동 스캔 (이름이 조금 달라도 유연하게 찾음)
        type_col = next((c for c in df.columns if any(k in str(c) for k in ['공종', '분야', '시설물', '종류', '구분'])), None)
        days_col = next((c for c in df.columns if any(k in str(c) for k in ['일수', '참여일', '기간', '인정'])), None)
        amt_col = next((c for c in df.columns if any(k in str(c) for k in ['금액', '백만원', '준공'])), None)
        
        total_days = 0.0
        total_count = 0.0
        total_amt = 0.0
        
        for _, row in df.iterrows():
            # 빈 행 스킵
            if (days_col and pd.isna(row[days_col])) and (amt_col and pd.isna(row[amt_col])):
                continue
                
            # 가중치 판별
            weight = 0.8 # 기본 기타 토목 80%
            if type_col and pd.notna(row[type_col]):
                val_type = str(row[type_col]).replace(' ', '')
                if '교량' in val_type or '터널' in val_type: weight = 1.0
                
            # 참여일수 합산
            if days_col and pd.notna(row[days_col]):
                try: total_days += float(row[days_col]) * weight
                except: pass
                
            # 금액 합산
            if amt_col and pd.notna(row[amt_col]):
                try: total_amt += float(row[amt_col]) * weight
                except: pass
                
            # 💡 실적 건수 = 조건에 맞는 행의 개수 (가중치 적용)
            total_count += (1 * weight)
            
        years = total_days / 365.0
        summary += f" - 환산 경력(년): {years:.2f}년\n"
        summary += f" - 환산 실적 건수: {total_count:.1f}건\n"
        summary += f" - 환산 실적 금액: {total_amt:,.0f}\n\n"
        
    return summary

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

    def write_cell_safe(self, sheet, r, c, val):
        cell = sheet.cell(row=r, column=c)
        if type(cell).__name__ == 'MergedCell':
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    sheet.cell(row=merged_range.min_row, column=merged_range.min_col).value = val
                    break
        else:
            cell.value = val

    def run_ai_dreamteam_optimizer(self, notice_text, excel_bytes, semi_fixed, db_summary):
        excel_structure = self.parse_excel_structure(excel_bytes)
        excel_json_str = json.dumps(excel_structure, ensure_ascii=False, indent=2)
        
        prompt = f"""
        당신은 어떠한 수학적 연산도 하지 않는 '좌표 매칭 로봇'입니다.
        [파이썬 시스템 자동 산출 팩트 데이터]에 명시된 환산 수치를 100% 맹신하여, [세부평가기준]의 배점 구간을 찾고 [자기평가표 엑셀 구조]의 정확한 행(row_num)에 <점수>와 <팩트>를 기입하세요.

        [★★★★★ 절대 준수 규칙]
        1. **계산 창조 금지:** 스스로 더하거나 나누지 마세요. 주어진 팩트 데이터(예: 윤석순 5.33년, 22건 등)를 그대로 적용하세요.
        2. **기술인 역할 매칭:** 엑셀 구조의 '사업책임'에는 PM의 데이터, '분야별책임'에는 PE, '분야별참여'에는 PES를 연결하세요.
        3. **신용평가등급:** 당사는 '{semi_fixed.get('credit_rating')}'입니다. 회사채 기준 세부기준표에서 BB- 등급은 **2.8점**입니다.
        4. **업무중첩도:** '{semi_fixed.get('overlap_level')}' 이므로 최고 감점 구간인 **3.6점(사책), 2.4점(분책)**을 적용하세요.
        5. **산출근거(reason) 초간결화:** 수식을 뺀 최종 숫자만 적으세요. (예: "5.33년", "22건", "1,288 백만원", "특급", "BB-", "350% 이상", "7.0%")
        6. **소계/총계 연산:** 엑셀 상의 '계', '소계', '총계' 행은 도출된 하위 항목 획득점수를 더해서 적고, 'reason'은 "-" 로 하세요.

        [사용자 확인 반고정 항목]
        - 신용평가등급: {semi_fixed.get('credit_rating')}
        - 신규고용율: {semi_fixed.get('new_hire_rate')}%
        - 업무중첩도 구간: {semi_fixed.get('overlap_level')}
        - 기술개발비율: {semi_fixed.get('investment_ratio')}%

        [세부평가기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 구조 (이 번호에 맞춰 점수 기입)]
        {excel_json_str}

        {db_summary}

        오직 순수 JSON으로만 반환:
        {{
            "row_results": [
                {{"row_num": 1, "score": "<찾은 총점>", "reason": "-"}},
                {{"row_num": 8, "score": "<찾은 점수>", "reason": "<환산 경력>"}}
            ],
            "pm": ["<사업책임 이름>"],
            "pe": ["<분야별책임 이름1>"],
            "pes": ["<분야별참여 이름1>"]
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
st.caption("※ 시트 자동 분류 및 행 단위(Row-count) 실적 스캔 완전체 엔진")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. DB/서류 관리", "⚙️ 2. 공고문 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
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
                
    with col2:
        st.subheader("Zone B: '기술인' 서류 스캔")
        if st.button("🔍 드라이브 재스캔"):
            scan_drive_archive_cached.clear()
            load_master_db_from_drive.clear()
            st.rerun()
        archive_data = scan_drive_archive_cached()
        if archive_data: st.success(f"📂 기술자 {len(archive_data)}명 스캔 완료!")

with tab2:
    st.markdown("### 📊 파이썬 자동 연산 모니터 (마스터 DB 팩트 체크)")
    st.caption("시스템이 마스터 DB의 각 시트(기술인)를 순회하며, 행 개수(건수)와 일수, 금액을 가중치(1.0/0.8)를 적용해 선행 계산한 결과입니다.")
    
    db_dict = load_master_db_from_drive()
    if db_dict:
        # DB 서머리 캐싱
        st.session_state.db_summary_cache = calculate_fact_data(db_dict)
        st.text_area("마스터 DB 자동 분석 결과", st.session_state.db_summary_cache, height=300)
        st.success("✅ 불필요한 매핑 없이 시트 이름과 행 개수, 컬럼 키워드 스캔을 통해 완벽한 팩트 데이터를 추출했습니다.")
    else:
        st.warning("구글 드라이브에 '마스터'가 포함된 엑셀 파일을 올려주세요.")

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

    if st.button("🚀 자기평가표 점수 자동 기입", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        if not st.session_state.raw_excel_bytes:
            st.error("엑셀 템플릿이 없습니다.")
        elif not st.session_state.db_summary_cache:
            st.error("마스터 DB 분석 결과가 없습니다. Tab 2를 확인하세요.")
        else:
            with st.spinner('시스템이 수학 계산을 완료하고 엑셀 원본 파일에 좌표를 찍어 덮어쓰고 있습니다...'):
                final_excel, log_df, pm, pe, pes = engine.run_ai_dreamteam_optimizer(
                    st.session_state.notice_text, 
                    st.session_state.raw_excel_bytes, 
                    st.session_state.semi_fixed, 
                    st.session_state.db_summary_cache
                )
                if final_excel:
                    st.success("🎉 산출 완료! Tab 4에서 완성된 원본 엑셀 파일을 다운로드하세요.")
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
