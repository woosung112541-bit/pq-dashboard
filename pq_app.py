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
if 'df_template' not in st.session_state: st.session_state.df_template = pd.DataFrame()
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
# 🧠 [Backend Engine] 좌표 주입 및 CoT (Chain-of-Thought) 엔진
# ==========================================
class PQScoringEngine:
    def run_ai_dreamteam_optimizer(self, notice_text, df_template, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("마스터 DB 엑셀 파일을 찾을 수 없습니다.")
            return pd.DataFrame(), [], [], []
            
        db_csv = master_db.to_csv(index=False)
        
        # 엑셀 각 행에 row_index를 부여하여 AI에게 전달 (구조 강제 고정)
        records = []
        for idx, row in df_template.iterrows():
            record = {"row_index": str(idx)}
            record.update(row.to_dict())
            records.append(record)
        template_json_str = json.dumps(records, ensure_ascii=False, indent=2)
        
        prompt = f"""
        당신은 건설엔지니어링 PQ '데이터 정밀 채점 매크로'입니다.
        [지정공고문]과 [세부평가기준]을 적용하여 [마스터 DB]의 데이터를 철저히 계산한 뒤, [자기평가표 엑셀 양식]의 빈칸을 채우세요.

        [★★★★★ 절대 준수 룰]
        1. **지정공고문 우선의 원칙:** [지정공고문]과 [세부평가기준]의 내용이 충돌할 경우 반드시 [지정공고문]의 내용을 우선하여 적용하세요.
        2. **엑셀 원본 행(Row) 유지:** 제공된 [엑셀 양식]의 row_index를 절대 변경하거나 누락하지 마세요. 소계, 총계 행도 반드시 계산해서 넣으세요.
        3. **만점 환각(Hallucination) 절대 금지:** 마스터 DB에 있는 날짜, 건수, 금액을 추출하여 반드시 step_by_step_calc에서 수학적 계산을 수행하세요. 
           - 경력(년) = 총 참여일수 / 365
           - 실적 가중치 = 교량/터널 100%, 기타 토목 80% 
           눈대중으로 '만점'이나 임의의 수치(예: 4.85년)를 적어내면 안 됩니다.
        4. **신용평가등급 강제 적용:** 당사의 신용평가등급은 '{semi_fixed.get('credit_rating')}'입니다. 세부기준표에서 이 등급이 속하는 구간(예: B-이상~BBB-미만 구간인 2.8점 등)을 정확히 찾아 감점 반영하세요.
        5. **데이터 부재시 0점:** 데이터가 부족하여 산출이 불가능한 항목은 획득점수에 0을 넣고, 근거에 "직접 입력 필요"라고 적으세요.
        6. **산출근거(reason) 극강의 간결성:** reason 필드에는 장황한 설명을 빼고, **최종 도출된 팩트 수치만 아주 짧게** 적으세요.
           (작성 예시: "특급", "5.33년", "22건", "1,288 백만원", "88 건", "해당없음", "BB-", "건설기술특허 5건", "350% 이상", "7%")
           소계나 총계 등 요약 행의 reason은 "-" 로 통일하세요.

        [사용자 확인 반고정 항목]
        - 신용평가등급: {semi_fixed.get('credit_rating')}
        - 신규고용율: {semi_fixed.get('new_hire_rate')}%
        - 업무중첩도 구간: {semi_fixed.get('overlap_level')}
        - 기술개발비율: {semi_fixed.get('investment_ratio')}%

        [지정공고문 및 세부기준 텍스트]
        {notice_text[:8000]} 

        [자기평가표 엑셀 양식 (이 데이터의 row_index 좌표를 Key로 사용할 것)]
        {template_json_str}

        [마스터 DB]
        {db_csv}

        오직 순수 JSON으로만 반환하세요. 각 row_index마다 "step_by_step_calc"를 통해 팩트 기반 수학적 계산을 거친 후 "score"와 "reason"을 도출하세요.
        {{
            "row_results": {{
                "0": {{"step_by_step_calc": "빈칸이므로 건너뜀", "score": "", "reason": ""}},
                "1": {{"step_by_step_calc": "전체 합계 도출: 39.9+25+9.8+10+6+0.3 = 91", "score": "91", "reason": "-"}},
                "4": {{"step_by_step_calc": "DB에서 김진규의 등급 확인 결과 특급. 배점 2점 만점 충족", "score": "2", "reason": "특급"}},
                "7": {{"step_by_step_calc": "윤석순 참여일수 교량 1000일*1.0 + 기타 500일*0.8 = 1400일. 1400/365 = 3.83년. 3년이상~4년미만 구간이므로 4.8점", "score": "4.8", "reason": "3.83년"}}
            }},
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
            
            ai_results = parsed.get("row_results", {})
            result_df = df_template.copy()
            
            # 엑셀 원본의 '자기평가' 및 '산출근거' 컬럼 찾기 (동적 매칭)
            score_col, reason_col = None, None
            for col in result_df.columns:
                col_str = str(col).replace('\n', '').replace(' ', '')
                if '자기평가' in col_str or '점수' in col_str: score_col = col
                elif '산출근거' in col_str or '산출' in col_str: reason_col = col
            
            # 매칭 실패 시 임의로 컬럼 추가
            if not score_col: score_col = "자기평가 점수"; result_df[score_col] = ""
            if not reason_col: reason_col = "산출근거"; result_df[reason_col] = ""
                
            for idx in result_df.index:
                idx_str = str(idx)
                if idx_str in ai_results:
                    val = ai_results[idx_str].get("score", "")
                    reason = ai_results[idx_str].get("reason", "")
                    
                    if val is None: val = ""
                    if reason is None: reason = ""
                    
                    # 헤더(점수)와 같은 기존 텍스트는 덮어쓰지 않고 보존
                    orig_val = str(result_df.at[idx, score_col]).strip()
                    if orig_val == "점수" or orig_val == "자기평가":
                        continue
                        
                    result_df.at[idx, score_col] = val
                    result_df.at[idx, reason_col] = reason
            
            return result_df, parsed.get("pm", []), parsed.get("pe", []), parsed.get("pes", [])
        except Exception as e:
            st.error(f"연산 오류: {e}")
            return pd.DataFrame(), [], [], []

engine = PQScoringEngine()

# ==========================================
# 🖥️ [Frontend]
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 엑셀 원본 구조 완벽 보존 (좌표 주입 & CoT 연산 엔진)")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 시뮬레이션", "🖨️ 4. 서류 패키징"])

# --- [Tab 1] ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고/양식 업로드")
        notice_files = st.file_uploader("PDF 공고문 및 Excel 자기평가표 업로드", type=['pdf', 'xlsx'], accept_multiple_files=True)
        if notice_files and api_key and st.button("🧠 업로드 문서 AI 분석", type="primary"):
            with st.spinner("엑셀 뼈대를 고정하고 공고문 기준을 파악 중입니다..."):
                notice_temp = ""
                df_ex_temp = pd.DataFrame()
                for file in notice_files:
                    if file.name.lower().endswith('.pdf'):
                        pdf = PyPDF2.PdfReader(file)
                        for page in pdf.pages[:20]: notice_temp += page.extract_text() or ""
                    elif file.name.lower().endswith('.xlsx'):
                        df_ex_temp = pd.read_excel(file).fillna("") # NaN을 빈칸으로 처리
                
                if not df_ex_temp.empty:
                    st.session_state.notice_text = notice_temp
                    st.session_state.df_template = df_ex_temp
                    st.success("✅ 문서 분석 완료! 엑셀 원본 구조가 시스템에 완벽히 고정되었습니다.")
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
    if not st.session_state.df_template.empty:
        st.write("- **업로드된 자기평가서 엑셀 뼈대 (요약/수정 절대 불가 상태):**")
        st.dataframe(st.session_state.df_template.astype(str), use_container_width=True)
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

    if st.button("🚀 자기평가서 점수 산출 (엑셀 좌표 주입)", type="primary", disabled=not st.session_state.semi_fixed_confirmed):
        if st.session_state.df_template.empty:
            st.error("엑셀 템플릿이 없습니다. Tab 1에서 엑셀을 업로드하세요.")
        else:
            with st.spinner('원본 엑셀의 줄(Row) 번호에 맞춰 정확한 팩트 데이터를 끼워 넣고 있습니다... (수학 연산 중)'):
                df_res, pm, pe, pes = engine.run_ai_dreamteam_optimizer(st.session_state.notice_text, st.session_state.df_template, st.session_state.semi_fixed)
                if not df_res.empty:
                    st.success("🎉 산출 완료! 엑셀 원본 구조가 100% 보존되었습니다.")
                    
                    def highlight_manual_input(row):
                        # 산출근거 컬럼(마지막 컬럼)에 직접입력 필요가 있으면 하이라이트
                        if len(row) > 0 and '직접 입력' in str(row.iloc[-1]):
                            return ['background-color: #ffe6e6'] * len(row)
                        return [''] * len(row)
                        
                    st.dataframe(df_res.astype(str).style.apply(highlight_manual_input, axis=1), use_container_width=True)
                    st.session_state.final_pq_score_table = df_res
                    st.session_state.dream_team = pm + pe + pes
                    st.info(f"👉 **최종 선발 명단:** {', '.join(st.session_state.dream_team)}")

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
