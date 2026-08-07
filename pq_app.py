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
        "pm_cnt": 1, "pe_cnt": 2, "pes_cnt": 2,
        "extra_settings": {}
    }
if 'semi_fixed' not in st.session_state:
    # 회사 단위로 "거의 고정"이지만 시점에 따라 바뀔 수 있는 값들.
    # 과거에는 신용등급(BB-) 하나만 프롬프트에 하드코딩했으나, 매 실행마다
    # 사용자가 직접 확인하고 넘어가도록 UI 입력값으로 전환한다.
    st.session_state.semi_fixed = {
        "credit_rating": "BB-",       # 재정상태 건실도 산정 기준
        "penalty_points": "해당없음",   # 참여업체·기술인 벌점
        "new_hire_rate": 0.0,          # 건설기술인 신규고용율 (%)
        "overlap_level": "최저",        # 업무중첩도 구간: 최저/중간/최고 중 선택
        "investment_ratio": 0.0,       # 기술개발투자실적 비율 (%)
        "patent_tech_count": 0,        # 건설신기술·특허 활용 건수
        "bid_restriction": "해당없음",  # 입찰참가제한·업무정지
        "inspection_penalty": "해당없음",  # 점검·진단 실시결과 처분 이력
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
            q="name contains '마스터' and trashed=false",
            fields="files(id, name, mimeType)"
        ).execute()
        items = results.get('files', [])
        
        if not items:
            load_master_db_from_drive.clear()
            return pd.DataFrame()
        
        file_id = items[0]['id']
        mime_type = items[0].get('mimeType', '')
        
        if mime_type == 'application/vnd.google-apps.spreadsheet':
            request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            request = drive_service.files().get_media(fileId=file_id)
            
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False: _, done = downloader.next_chunk()
        fh.seek(0)
        
        df = pd.read_excel(fh)
        if df.empty: load_master_db_from_drive.clear()
        return df
        
    except Exception:
        load_master_db_from_drive.clear()
        return pd.DataFrame()

def get_all_files_recursively(drive_service, folder_id, target_mime_types, depth=0):
    if depth > 5: return [] 
    found_files = []
    query = f"'{folder_id}' in parents and trashed=false"
    page_token = None
    while True:
        try:
            response = drive_service.files().list(q=query, fields='nextPageToken, files(id, name, mimeType)', pageToken=page_token).execute()
            for file in response.get('files', []):
                if file['mimeType'] in target_mime_types:
                    found_files.append(file)
                elif file['mimeType'] == 'application/vnd.google-apps.folder':
                    found_files.extend(get_all_files_recursively(drive_service, file['id'], target_mime_types, depth + 1))
            page_token = response.get('nextPageToken', None)
            if not page_token: break
        except Exception: break
    return found_files

def get_all_subfolders_map(drive_service, root_id, depth=0):
    if depth > 5: return {}
    folder_dict = {}
    query = f"'{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    page_token = None
    while True:
        try:
            response = drive_service.files().list(q=query, fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
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
        q_arch = "name='기술인' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_arch = drive_service.files().list(q=q_arch, fields="files(id)").execute()
        if not res_arch.get('files'): return {}
        
        archive_id = res_arch.get('files')[0]['id']
        q_sub = f"'{archive_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_sub = drive_service.files().list(q=q_sub, fields="files(id, name)").execute()
        
        archive_status = {}
        for folder in res_sub.get('files', []):
            all_pdfs = get_all_files_recursively(drive_service, folder['id'], ['application/pdf'])
            archive_status[folder['name']] = [f['name'] for f in all_pdfs]
        return archive_status
    except Exception: return {"시스템 스캔 오류 (새로고침을 눌러주세요)": []}

def get_target_project_folders(drive_service):
    try:
        q = "name='작성대상' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res = drive_service.files().list(q=q, fields="files(id)").execute()
        items = res.get('files', [])
        if not items: return {}
        
        target_id = items[0]['id']
        q_sub = f"'{target_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_sub = drive_service.files().list(q=q_sub, fields="files(id, name)").execute()
        return {f['name']: f['id'] for f in res_sub.get('files', [])}
    except Exception: return {}

def get_ai_model():
    return genai.GenerativeModel('gemini-3.6-flash')

# ==========================================
# 🧠 [Backend Engine] AI 다이렉트 100점 만점 시뮬레이션 엔진
# ==========================================
class PQScoringEngine:
    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt, notice_text, self_eval_template, semi_fixed):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("❌ 구글 드라이브에서 '마스터'라는 이름이 포함된 엑셀 파일을 찾을 수 없습니다.")
            return pd.DataFrame(), [], [], []
            
        db_csv = master_db.to_csv(index=False)
        available_engineers = list(scan_drive_archive_cached().keys())
        engineers_str = ", ".join(available_engineers) if available_engineers else "명단 없음"
        
        prompt = f"""
        당신은 건설엔지니어링 PQ 최고 심사위원입니다.
        아래 제공된 [공고문 텍스트(PDF)], [자기평가표 양식(Excel)], [엔지니어 실적 Master DB]를 융합 분석하여 최적의 드림팀과 평가 점수표를 산출하세요.

        [★★★★★ 절대 준수 규칙]
        1. **완벽한 서식 미러링:** 반환하는 JSON "pq_score_table"은 아래 제공된 [자기평가표 양식(Excel)]의 행/열 구조와 배점을 토씨 하나 틀리지 말고 100% 똑같이 유지하세요. 당신은 획득점수와 계산근거만 채웁니다.
        2. **데이터 부재 시 "직접 입력 필요":** 마스터 DB 및 [기술인] 드라이브 폴더 서류를 조회했을 때 실적 데이터(백파일)가 없거나 산출 근거를 도저히 계산할 수 없는 항목은 억지로 점수를 부여하지 마세요! 획득점수를 빈칸(또는 0)으로 두고, 점수계산근거에 반드시 **"직접 입력 필요"**라고 적으세요.
        3. **만점 환각(Hallucination) 금지:** DB에 실적이 있다면 반드시 '실제 건수/금액'을 바탕으로 공고문 및 세부평가기준 상의 제한사항(최근 N년, 준공 여부, 직무분야=토목 한정, 자체안전점검 제외, 교량·터널 100%/기타토목 80% 가중치 등)을 정확히 적용해 깎을 건 깎으세요. 근거에 실제 수치와 적용한 제한조건을 적으세요.
        4. **사용자 확인 반고정값 그대로 사용:** 아래 [사용자 확인 반고정 항목]은 회사 단위로 거의 고정이지만 시점마다 바뀔 수 있어 매번 사용자가 직접 확인한 값입니다. 당신이 임의로 추정하지 말고 주어진 값을 그대로 적용하세요.
        5. **등급·경력·실적·유사용역수행실적은 자료 기반 자동 산출:** 위 반고정 항목에 포함되지 않은 모든 항목(등급, 경력, 실적건수/금액, 유사용역수행실적 등)은 마스터 DB와 [기술인] 드라이브 폴더의 실제 서류만 근거로 계산하세요. 유사용역수행실적은 참여기술인 개인별 실적을 합산하는 것이 아니라, 동일 계약이 여러 기술인 시트에 중복 기재된 경우 회사 단위로 한 번만 인정해야 합니다.

        [사용자 확인 반고정 항목]
        - 신용평가등급: {semi_fixed.get('credit_rating')}
        - 참여업체·기술인 벌점: {semi_fixed.get('penalty_points')}
        - 건설기술인 신규고용율: {semi_fixed.get('new_hire_rate')}%
        - 업무중첩도 구간: {semi_fixed.get('overlap_level')}
        - 기술개발투자실적 비율: {semi_fixed.get('investment_ratio')}%
        - 건설신기술·특허 활용 건수: {semi_fixed.get('patent_tech_count')}건
        - 입찰참가제한·업무정지: {semi_fixed.get('bid_restriction')}
        - 점검·진단 실시결과 처분 이력: {semi_fixed.get('inspection_penalty')}

        [후보군 및 조건]
        - 가용 기술자 명단: [{engineers_str}]
        - 필요 인원: 사책 {pm_cnt}명, 분책 {pe_cnt}명, 분참 {pes_cnt}명 (기간은 2026년 기준 3년 환산)

        [공고문 전체 텍스트 (세부기준)]
        {notice_text[:7000]} 

        [자기평가표 양식 (Excel 템플릿 내용)]
        {self_eval_template}

        [엔지니어 실적 Master DB (CSV 형식)]
        {db_csv}

        분석 결과를 오직 아래 JSON 포맷으로만 반환하세요.
        {{
            "pq_score_table": [
                {{"분류": "...", "세부평가항목": "...", "배점": "...", "획득점수": 0, "점수계산근거": "직접 입력 필요"}},
                {{"분류": "...", "세부평가항목": "...", "배점": "...", "획득점수": 4.5, "점수계산근거": "건수 5건 확인 적용"}},
                {{"분류": "신용도", "세부평가항목": "재정상태 건실도", "배점": 3, "획득점수": 1.5, "점수계산근거": "신용등급 BB- 적용"}}
            ],
            "pm": ["선발된 이름1"],
            "pe": ["선발된 이름2", "선발된 이름3"],
            "pes": ["선발된 이름4", "선발된 이름5"]
        }}
        """
        
        try:
            model = get_ai_model()
            response = model.generate_content(prompt)
            result_text = response.text.strip().removeprefix("```json").removesuffix("```").strip()
            parsed = json.loads(result_text)
            
            df = pd.DataFrame(parsed.get("pq_score_table", []))
            return df, parsed.get("pm", []), parsed.get("pe", []), parsed.get("pes", [])
        except Exception as e:
            st.error(f"AI 연산 중 오류 발생: {e}")
            return pd.DataFrame(), [], [], []

engine = PQScoringEngine()

# ==========================================
# 🖥️ [Frontend] 메인 대시보드 UI
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 엑셀(자기평가표) 완벽 미러링 지원 및 '직접 입력 필요' 검증 탑재")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 책임기술자 시뮬레이션", "🖨️ 4. 서류 출력 및 패키징"])

# --- [Tab 1] 마스터 DB 및 드라이브 스캔 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        
        upload_method = st.radio("입력 방식", ["☁️ 구글 드라이브 '작성대상' 폴더에서 선택", "📤 직접 드래그 앤 드롭 (PDF + Excel)"], horizontal=True)
        
        notice_text_temp = ""
        self_eval_template_temp = ""
        ready_to_analyze = False
        
        if upload_method == "📤 직접 드래그 앤 드롭 (PDF + Excel)":
            notice_files = st.file_uploader("안내서/세부기준(PDF) 및 자기평가표 양식(Excel) 업로드", type=['pdf', 'xlsx', 'xls'], accept_multiple_files=True)
            if notice_files and api_key and st.button("🧠 업로드한 공고/엑셀 AI 분석", type="primary"):
                with st.spinner("문서를 읽고 엑셀 서식을 파악하는 중..."):
                    for file in notice_files:
                        if file.name.lower().endswith('.pdf'):
                            pdf = PyPDF2.PdfReader(file)
                            for page in pdf.pages[:15]: notice_text_temp += page.extract_text() or ""
                        elif file.name.lower().endswith(('.xlsx', '.xls')):
                            df_excel = pd.read_excel(file)
                            self_eval_template_temp += f"\n--- [{file.name}] ---\n" + df_excel.to_csv(index=False)
                    ready_to_analyze = True
                    
        else: # ☁️ 드라이브 선택 모드
            if api_key:
                drive_service = authenticate_google_drive()
                if drive_service:
                    project_folders = get_target_project_folders(drive_service)
                    if not project_folders:
                        st.info("💡 구글 드라이브에 `작성대상` 폴더가 없거나 비어있습니다.")
                    else:
                        selected_project_name = st.selectbox("분석할 공고(프로젝트) 폴더를 선택하세요:", list(project_folders.keys()))
                        if st.button("🧠 선택한 공고 AI 분석", type="primary"):
                            with st.spinner(f"폴더 안의 PDF 및 Excel 양식을 모두 읽어오는 중..."):
                                target_id = project_folders[selected_project_name]
                                target_mime_types = [
                                    'application/pdf', 
                                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                    'application/vnd.ms-excel',
                                    'application/vnd.google-apps.spreadsheet'
                                ]
                                doc_files = get_all_files_recursively(drive_service, target_id, target_mime_types, depth=1)
                                
                                if not doc_files:
                                    st.warning("선택한 폴더 안에 문서가 없습니다.")
                                else:
                                    for doc in doc_files:
                                        if doc['mimeType'] == 'application/pdf':
                                            request = drive_service.files().get_media(fileId=doc['id'])
                                            fh = io.BytesIO()
                                            downloader = MediaIoBaseDownload(fh, request)
                                            done = False
                                            while not done: _, done = downloader.next_chunk()
                                            fh.seek(0)
                                            pdf = PyPDF2.PdfReader(fh)
                                            for page in pdf.pages[:15]: notice_text_temp += page.extract_text() or ""
                                        else:
                                            if doc['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                                                request = drive_service.files().export_media(fileId=doc['id'], mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                                            else:
                                                request = drive_service.files().get_media(fileId=doc['id'])
                                            fh = io.BytesIO()
                                            downloader = MediaIoBaseDownload(fh, request)
                                            done = False
                                            while not done: _, done = downloader.next_chunk()
                                            fh.seek(0)
                                            df_excel = pd.read_excel(fh)
                                            self_eval_template_temp += f"\n--- [{doc['name']}] ---\n" + df_excel.to_csv(index=False)
                                    ready_to_analyze = True
        
        if ready_to_analyze:
            with st.spinner("AI가 공고문을 정독하며 세부기준을 엑셀 양식과 매칭 중입니다..."):
                try:
                    st.session_state.notice_text = notice_text_temp
                    st.session_state.self_eval_template = self_eval_template_temp
                    
                    prompt = f"""
                    건설엔지니어링 PQ 텍스트/엑셀 분석 후 순수 JSON만 반환.
                    1. eval_criteria: 배점표 배열 (대략적인 요약)
                    2. settings: {{ has_safety(bool), period(str), bohal(list), pm_cnt(int), pe_cnt(int), pes_cnt(int), extra_settings: {{"특이사항명": "내용"}} }}
                    공고문 텍스트: {notice_text_temp}
                    자기평가서 엑셀 구조: {self_eval_template_temp}
                    """
                    response = get_ai_model().generate_content(prompt)
                    parsed_json = json.loads(response.text.strip().removeprefix("```json").removesuffix("```").strip())
                    
                    st.session_state.eval_criteria = pd.DataFrame(parsed_json.get("eval_criteria", []))
                    st.session_state.auto_settings = parsed_json.get("settings", st.session_state.auto_settings)
                    st.success("✅ 분석 완료! 엑셀 양식 구조와 세부 기준이 메모리에 완벽히 저장되었습니다.")
                except Exception as e:
                    st.error(f"분석 실패: {e}")
                        
    with col2:
        st.subheader("Zone B: '기술인' 폴더 실적 아카이브 현황")
        st.info("💡 실적증명서 및 수료증 PDF는 구글 드라이브 `기술인` 폴더 내에 업로드하시면 됩니다.")
        if st.button("🔍 구글 드라이브 아카이브 현황 새로고침"):
            scan_drive_archive_cached.clear()
            load_master_db_from_drive.clear()
            st.rerun()

        with st.spinner("구글 드라이브 '기술인' 폴더 딥-스캔 중..."):
            archive_data = scan_drive_archive_cached()
            if archive_data:
                st.success(f"📂 기술자 총 {len(archive_data)}명의 폴더 스캔 완료!")
                for name, pdfs in archive_data.items():
                    with st.expander(f"📁 **{name}** (총 {len(pdfs)}개 서류 보관 중)"):
                        if pdfs:
                            for pdf in pdfs: st.write(f"- 📄 `{pdf}`")
                        else:
                            st.caption("해당 폴더 내부에는 어떠한 PDF도 존재하지 않습니다.")
            else:
                st.warning("구글 드라이브 최상단에 `기술인` 폴더가 없거나 비어 있습니다.")

# --- [Tab 2] 공고문 세부사항 설정 ---
with tab2:
    st.markdown("### 📊 공고문 AI 분석 결과 및 세부 설정")
    col_title, col_toggle = st.columns([7, 3])
    with col_toggle: manual_override = st.toggle("⚙️ 세부사항 수동 설정", value=False)

    if not st.session_state.eval_criteria.empty: st.table(st.session_state.eval_criteria)
    else: st.info("💡 [Tab 1]에서 공고문/엑셀을 업로드하시면 평가 배점표가 자동으로 구성됩니다.")
    
    st.markdown("---")
    s_settings = st.session_state.auto_settings
    
    bohal_data = s_settings.get('bohal', [])
    if not isinstance(bohal_data, list) or len(bohal_data) == 0: 
        bohal_data = [{"전문분야": "해당없음", "비율(%)": 0}]

    extra_settings = s_settings.get('extra_settings', {})
    if not isinstance(extra_settings, dict): extra_settings = {"기타사항": str(extra_settings)}

    if not manual_override:
        st.success("🤖 AI 자동 세팅 모드입니다.")
        col_a, col_b = st.columns(2)
        with col_a:
            st.write(f"- **정기안전점검 포함:** {'✅' if s_settings.get('has_safety', False) else '❌'}")
            st.write(f"- **실적 인정 기간:** {s_settings.get('period', '제한없음')}")
            st.write(f"- **필요 인원:** 사책 {s_settings.get('pm_cnt', 1)} / 분책 {s_settings.get('pe_cnt', 0)} / 분참 {s_settings.get('pes_cnt', 0)}")
            st.write("- **기업 신용평가등급:** **BB- (절대 고정/감점 적용)**")
        with col_b:
            st.write("- **보할 인정 비율:**")
            st.table(pd.DataFrame(bohal_data))
            
        final_pm_cnt, final_pe_cnt, final_pes_cnt = s_settings.get('pm_cnt', 1), s_settings.get('pe_cnt', 0), s_settings.get('pes_cnt', 0)
        
        st.markdown("##### 📌 공고문 특이/세부사항 (AI 자동 추출)")
        if extra_settings:
            for k, v in extra_settings.items(): st.info(f"**{k}** : {v}")
        else: st.caption("별도의 특이사항이 발견되지 않았습니다.")
            
    else:
        st.warning("⚠️ 수동 설정 모드입니다.")
        chk_safety = st.checkbox("✅ 정기안전점검 포함", value=s_settings.get('has_safety', True))
        sel_period = st.selectbox("↳ 인정 기간", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
        st.write("**✅ 보할 설정**")
        edited_bohal = st.data_editor(pd.DataFrame(bohal_data), num_rows="dynamic")
        col_pm, col_pe, col_pes = st.columns(3)
        with col_pm: final_pm_cnt = st.number_input("사책(명)", value=s_settings.get('pm_cnt', 1))
        with col_pe: final_pe_cnt = st.number_input("분책(명)", value=s_settings.get('pe_cnt', 0))
        with col_pes: final_pes_cnt = st.number_input("분참(명)", value=s_settings.get('pes_cnt', 0))
        st.markdown("##### 📌 특이/세부사항 수동 편집")
        extra_df = pd.DataFrame(list(extra_settings.items()), columns=["항목명", "내용"])
        edited_extra = st.data_editor(extra_df, num_rows="dynamic", use_container_width=True)

# --- [Tab 3] 책임기술자 시뮬레이션 결과 ---
with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 및 엄격한 점수 산출")

    CREDIT_GRADES = ["AAA","AA+","AA0","AA-","A+","A0","A-","BBB+","BBB0","BBB-",
                      "BB+","BB0","BB-","B+","B0","B-","CCC+ 이하","미제출"]

    with st.expander("🔒 반고정 항목 확인 (실행 전 필수)", expanded=not st.session_state.semi_fixed_confirmed):
        st.caption("회사 단위로 거의 고정이지만 시점에 따라 바뀔 수 있는 값입니다. AI가 임의로 추정하지 않고, 아래에서 직접 확인한 값을 그대로 채점에 반영합니다.")
        sf = st.session_state.semi_fixed
        col1, col2 = st.columns(2)
        with col1:
            sf['credit_rating'] = st.selectbox("신용평가등급", CREDIT_GRADES,
                                                index=CREDIT_GRADES.index(sf['credit_rating']) if sf['credit_rating'] in CREDIT_GRADES else 0)
            sf['penalty_points'] = st.text_input("참여업체·기술인 벌점 (최근 2년 누계평균, 없으면 '해당없음')", value=sf['penalty_points'])
            sf['new_hire_rate'] = st.number_input("건설기술인 신규고용율 (%)", value=float(sf['new_hire_rate']), step=0.1)
            sf['overlap_level'] = st.selectbox("업무중첩도 구간(사업책임·분야별책임)", ["최저","중간","최고"],
                                                index=["최저","중간","최고"].index(sf['overlap_level']))
        with col2:
            sf['investment_ratio'] = st.number_input("기술개발투자실적 비율 (%, 최근 3년 매출대비)", value=float(sf['investment_ratio']), step=0.01)
            sf['patent_tech_count'] = st.number_input("건설신기술·특허 활용 건수", value=int(sf['patent_tech_count']), step=1)
            sf['bid_restriction'] = st.text_input("입찰참가제한·업무정지 (없으면 '해당없음')", value=sf['bid_restriction'])
            sf['inspection_penalty'] = st.text_input("점검·진단 실시결과 처분 이력 (없으면 '해당없음')", value=sf['inspection_penalty'])
        st.session_state.semi_fixed = sf
        st.session_state.semi_fixed_confirmed = st.checkbox(
            "위 값을 직접 확인했습니다. 이번 실행에 그대로 사용합니다.", value=st.session_state.semi_fixed_confirmed)

    if not st.session_state.semi_fixed_confirmed:
        st.caption("⬆️ 반고정 항목을 확인 체크해야 시뮬레이션을 실행할 수 있습니다.")

    if st.button("🚀 마스터 DB 딥러닝 시뮬레이션 실행 (엑셀 자기평가표 미러링)", type="primary",
                  disabled=not st.session_state.semi_fixed_confirmed):
        if not st.session_state.notice_text or not st.session_state.self_eval_template:
            st.warning("⚠️ [Tab 1]에서 공고문(PDF)과 자기평가표(Excel)를 먼저 분석해야 합니다!")
        else:
            with st.spinner('AI가 엑셀 양식을 복제하고 데이터 부재 항목은 "직접 입력 필요"로 분류하며 연산 중입니다... (약 15초 소요)'):
                best_score_df, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(
                    final_pm_cnt, final_pe_cnt, final_pes_cnt, st.session_state.notice_text,
                    st.session_state.self_eval_template, st.session_state.semi_fixed
                )
                
                if not best_score_df.empty:
                    st.success("🎉 AI 엑셀 기반 자기평가서 실점수 산출 완료!")
                    
                    best_score_df['배점_num'] = pd.to_numeric(best_score_df['배점'], errors='coerce')
                    best_score_df['획득점수_num'] = pd.to_numeric(best_score_df['획득점수'], errors='coerce')
                    
                    total_allocated = best_score_df['배점_num'].sum()
                    total_earned = best_score_df['획득점수_num'].sum()
                    
                    col1, col2 = st.columns(2)
                    col1.metric("총 배점 (합계)", f"{total_allocated:g}점")
                    col2.metric("최종 획득 점수", f"{total_earned:g}점")
                    
                    display_df = best_score_df.drop(columns=['배점_num', '획득점수_num'])
                    
                    def highlight_manual_input(row):
                        if '직접 입력 필요' in str(row.get('점수계산근거', '')):
                            return ['background-color: #ffe6e6'] * len(row)
                        return [''] * len(row)
                        
                    st.dataframe(display_df.style.apply(highlight_manual_input, axis=1), use_container_width=True)
                    st.session_state.final_pq_score_table = display_df
                    
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
            with st.spinner("드라이브 '기술인' 폴더를 뒤져 해당 인원의 모든 증빙 서류를 추출 중입니다..."):
                try:
                    drive_service = authenticate_google_drive()
                    archive_query = "name='기술인' and mimeType='application/vnd.google-apps.folder' and trashed=false"
                    archive_res = drive_service.files().list(q=archive_query, fields="files(id)").execute()
                    
                    if not archive_res.get('files'):
                        st.error("구글 드라이브에 `기술인` 폴더가 존재하지 않습니다.")
                    else:
                        archive_id = archive_res['files'][0]['id']
                        zip_buffer = io.BytesIO()
                        found_files_count = 0
                        folder_map = get_all_subfolders_map(drive_service, archive_id)
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                            if not st.session_state.final_pq_score_table.empty:
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                    st.session_state.final_pq_score_table.to_excel(writer, index=False, sheet_name='자기평가서_산출표')
                                z.writestr("0_자기평가서_양식_산출표.xlsx", excel_buffer.getvalue())
                                
                            for person_name in st.session_state.dream_team:
                                if person_name in folder_map:
                                    person_folder_id = folder_map[person_name]
                                    person_pdfs = get_all_files_recursively(drive_service, person_folder_id, ['application/pdf'])
                                    
                                    if person_pdfs:
                                        for pdf_file in person_pdfs:
                                            try:
                                                request = drive_service.files().get_media(fileId=pdf_file['id'])
                                                fh = io.BytesIO()
                                                downloader = MediaIoBaseDownload(fh, request)
                                                done = False
                                                while not done: _, done = downloader.next_chunk()
                                                fh.seek(0)
                                                z.writestr(f"{person_name}/{pdf_file['name']}", fh.read())
                                                found_files_count += 1
                                            except Exception: pass
                                    else:
                                        z.writestr(f"{person_name}/안내_서류없음.txt", "폴더는 있으나 내부에 PDF 서류가 없습니다.".encode('utf-8'))
                                else:
                                    z.writestr(f"{person_name}/안내_폴더없음.txt", f"구글 드라이브 '기술인' 폴더에 '{person_name}' 폴더가 어디에도 없습니다.".encode('utf-8'))
                        
                        zip_buffer.seek(0)
                        if found_files_count > 0:
                            st.success(f"🎉 성공! 총 {found_files_count}개의 증빙 PDF와 자기평가서 산출표를 찾아 압축했습니다.")
                            st.download_button("📦 최종 제출서류 패키지 다운로드 (ZIP)", data=zip_buffer, file_name="최종_PQ제출서류_패키지.zip", mime="application/zip", type="primary")
                        else: st.warning("선발된 인원에 해당하는 PDF 서류를 찾지 못했습니다.")
                except Exception as e: st.error(f"서류 패키징 중 오류 발생: {e}")
