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
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# ⚙️ [초기 세팅] 시스템 가상 메모리 및 페이지 설정
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")

if 'uploaded_pdfs' not in st.session_state:
    st.session_state.uploaded_pdfs = {}
if 'eval_criteria' not in st.session_state:
    st.session_state.eval_criteria = pd.DataFrame() # Zone A에서 뽑아낼 배점표 저장 공간

# 사이드바: AI API 키 입력
with st.sidebar:
    st.markdown("### 🧠 AI 엔진 설정")
    api_key = st.text_input("Gemini API Key를 입력하세요", type="password")
    if api_key:
        genai.configure(api_key=api_key)
        st.success("AI 엔진 연결 완료!")
    else:
        st.warning("문서 자동 분석을 위해 API Key가 필요합니다.")

# ==========================================
# 🔗 [Data Loader] 구글 드라이브 마스터 DB 연동
# ==========================================
@st.cache_data(ttl=600)
def load_master_db_from_drive():
    try:
        if "GOOGLE_CREDENTIALS" not in st.secrets:
            return pd.DataFrame()
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds = Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/drive']
        )
        drive_service = build('drive', 'v3', credentials=creds)

        results = drive_service.files().list(
            q="name contains '마스터' and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false",
            fields="files(id, name)"
        ).execute()
        items = results.get('files', [])

        if not items:
            return pd.DataFrame()

        file_id = items[0]['id']
        request = drive_service.files().get_media(fileId=file_id)
        
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        
        fh.seek(0)
        return pd.read_excel(fh)
        
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 🧠 [Backend Engine] PQ 점수 계산 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        self.master_db = load_master_db_from_drive()

    def get_personnel_list(self):
        if self.master_db.empty:
            return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]
        for col_name in ['이름', '성명', '엔지니어명', '기술자명', '기술인']:
            if col_name in self.master_db.columns:
                names = self.master_db[col_name].dropna().unique().tolist()
                return ["(선택)"] + names
        return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]

    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt):
        best_score_df = pd.DataFrame({
            "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
            "배점": [30, 20, 10, -5],
            "획득점수": [30.0, 20.0, 10.0, 0.0],
            "비고": ["AI 최적화", "중첩도 0건", "A+ 등급", "해당없음"]
        })
        sample_names = ["윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장"]
        pm_list = sample_names[0:pm_cnt] if pm_cnt > 0 else []
        pe_list = sample_names[pm_cnt:pm_cnt+pe_cnt] if pe_cnt > 0 else []
        pes_list = sample_names[pm_cnt+pe_cnt:pm_cnt+pe_cnt+pes_cnt] if pes_cnt > 0 else []
        return best_score_df, pm_list, pe_list, pes_list

    def calculate_manual_score(self):
        return pd.DataFrame({
            "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
            "배점": [30, 20, 10, -5],
            "획득점수": [28.5, 20.0, 10.0, 0.0],
            "비고": ["수동 선택 검증 완료", "이상 없음", "우수", "해당 없음"]
        })

engine = PQScoringEngine()

def get_ai_model():
    """안정성을 위해 3.6 시도 후 1.5로 우회하는 모델 호출 헬퍼 함수"""
    try:
        return genai.GenerativeModel('gemini-3.6-flash')
    except:
        return genai.GenerativeModel('gemini-1.5-flash')

# ==========================================
# 🖥️ [Frontend] 메인 대시보드 UI
# ==========================================
st.title("PQ 자동화 대시보드")
st.caption("※ 본 페이지는 로컬 및 클라우드 테스트용 프로토타입입니다.")

tab1, tab2, tab3, tab4 = st.tabs([
    "📥 1. 마스터 DB 관리", 
    "⚙️ 2. 공고문 세부사항 설정", 
    "📊 3. 책임기술자 시뮬레이션 결과", 
    "🖨️ 4. 서류 출력 및 패키징"
])

# --- [Tab 1] 마스터 DB 관리 ---
with tab1:
    col1, col2 = st.columns(2)
    
    # ---------------------------------------------------------
    # 👉 [기능 통합 1] Zone A: 공고문 AI 파싱 뼈대
    # ---------------------------------------------------------
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF)을 드래그 앤 드롭하세요.", type=['pdf'], accept_multiple_files=True, key="zone_a")
        
        if notice_files and api_key:
            if st.button("🧠 공고문 AI 분석 및 배점표 자동 생성", type="primary"):
                with st.spinner("AI가 공고문을 정독하며 자기평가 배점표를 추출하고 있습니다..."):
                    try:
                        # 1. PDF 텍스트 추출 (앞 5페이지만 읽어 시간/비용 절약)
                        notice_text = ""
                        for file in notice_files:
                            pdf = PyPDF2.PdfReader(file)
                            for page in pdf.pages[:5]:
                                notice_text += page.extract_text() or ""
                        
                        # 2. AI에게 배점표 추출 지시
                        prompt = f"""
                        다음은 건설엔지니어링(PQ) 공고문/지침서 내용입니다.
                        이 내용을 분석해서 평가 항목과 배점 기준을 파악한 뒤 JSON 배열 형태로 반환하세요.
                        배열의 각 항목은 다음 키를 가져야 합니다: "대분류", "평가항목", "배점", "세부인정기준"
                        
                        [공고문 내용]
                        {notice_text}
                        
                        반드시 순수 JSON 문자열만 출력하세요. 마크다운(```json)은 쓰지 마세요.
                        """
                        model = get_ai_model()
                        response = model.generate_content(prompt)
                        
                        # 응답 텍스트 정제
                        result_text = response.text.strip()
                        if result_text.startswith("```json"): result_text = result_text[7:-3].strip()
                        elif result_text.startswith("```"): result_text = result_text[3:-3].strip()
                        
                        parsed_json = json.loads(result_text)
                        st.session_state.eval_criteria = pd.DataFrame(parsed_json)
                        
                        st.success("✅ 공고문 분석 성공! [Tab 2]에 배점표가 세팅되었습니다.")
                    except Exception as e:
                        st.error(f"공고문 분석 중 에러 발생: {e}")
        elif notice_files and not api_key:
            st.error("👈 왼쪽 사이드바에 API Key를 먼저 입력해주세요!")
            
    # ---------------------------------------------------------
    # 👉 [기능 통합 2] Zone B: 실적 업데이트 및 AI 스마트 중복 회피
    # ---------------------------------------------------------
    with col2:
        st.subheader("Zone B: 실적 업데이트 (AI 스마트 스캔)")
        perf_file = st.file_uploader("기술인/회사 실적증명서(PDF) 업로드", type=['pdf'], key="zone_b")
        
        if perf_file:
            if not api_key:
                st.error("👈 왼쪽 사이드바에 Gemini API Key를 먼저 입력해주세요!")
            else:
                with st.spinner("🧠 AI가 기존 DB와 대조하며 '순수 신규 실적'만 추출 중입니다..."):
                    try:
                        pdf_part = {"mime_type": "application/pdf", "data": perf_file.getvalue()}
                        
                        # 기존 사업명 목록 텍스트로 준비
                        existing_projects_list = []
                        if not engine.master_db.empty and '사업명' in engine.master_db.columns:
                            existing_projects_list = engine.master_db['사업명'].dropna().tolist()
                        existing_str = ", ".join(map(str, existing_projects_list))
                        
                        # 👉 AI 스마트 중복 회피 프롬프트
                        prompt = f"""
                        이 PDF 문서를 분석하여 다음 정보를 JSON 형식으로만 반환해주세요.
                        1. doc_type: (경력증명서, 실적증명서, 신용평가등급확인서, 교육수료증, 기타증빙서류 중 하나)
                        2. owner: 기술자 이름 (없으면 '회사공통')
                        3. projects: 문서에 기재된 주요 사업 실적 목록 (배열). 키값: "사업명", "시작일", "종료일", "담당업무", "발주처"
                        
                        [🚨 매우 중요한 중복 회피 규칙 🚨]
                        현재 우리 회사 데이터베이스에 이미 등록된 사업명 목록은 다음과 같습니다:
                        [{existing_str}]
                        
                        PDF에서 찾은 사업 실적 중에서, 위 목록과 '의미상/내용상 같은 사업'이라고 판단되는 것(띄어쓰기, 괄호, 축약어 등의 미세한 차이 포함)은 'projects' 배열에서 무조건 완전히 제외시키세요.
                        오직 기존 DB에 절대 없는 '완벽한 신규 사업'만 'projects' 배열에 담아서 반환해야 합니다. 신규 사업이 아예 없으면 빈 배열 [] 을 반환하세요.
                        
                        반드시 순수 JSON 문자열만 출력하세요. 마크다운(```json)은 쓰지 마세요.
                        """
                        
                        model = get_ai_model()
                        response = model.generate_content([prompt, pdf_part])
                        
                        result_text = response.text.strip()
                        if result_text.startswith("```json"): result_text = result_text[7:-3].strip()
                        elif result_text.startswith("
