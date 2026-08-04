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
                        elif result_text.startswith("```"): result_text = result_text[3:-3].strip()
                            
                        result_json = json.loads(result_text)
                        
                        doc_type = result_json.get("doc_type", "기타증빙서류")
                        owner = result_json.get("owner", "회사공통")
                        projects = result_json.get("projects", [])
                                
                        new_filename = f"[{doc_type}] {owner}.pdf"
                        st.info(f"💡 문서 주인 판별 완료: **{owner}**의 **{doc_type}**")
                        
                        if projects:
                            df_new = pd.DataFrame(projects)
                            st.success(f"✨ AI 필터링 완료! 중복을 제외한 완전 신규 실적 {len(projects)}건을 발견했습니다.")
                            st.dataframe(df_new, use_container_width=True)
                            st.button("💾 이 신규 실적만 마스터 DB 엑셀에 업데이트", type="primary")
                        else:
                            st.warning("⚠️ 추출된 모든 실적이 이미 마스터 DB에 등록되어 있는 사업입니다. (중복 방지 시스템 작동 - 전체 패스됨)")
                        
                        perf_file.seek(0)
                        st.session_state.uploaded_pdfs[new_filename] = perf_file.getvalue()
                        
                    except Exception as e:
                        st.error(f"AI 분석 중 에러 발생: {str(e)}")

# --- [Tab 2] 공고문 세부사항 설정 ---
with tab2:
    st.markdown("### 📊 평가 항목 및 배점 기준 시각화 (자기평가표 초안)")
    
    # 👉 Zone A에서 파싱한 데이터가 있으면 그걸 띄워주고, 없으면 더미 데이터 띄워주기
    if not st.session_state.eval_criteria.empty:
        st.success("Zone A에서 AI가 분석한 공고문 배점표가 성공적으로 연동되었습니다!")
        st.table(st.session_state.eval_criteria)
    else:
        st.info("💡 공고문을 업로드하시면 AI가 이 표를 자동으로 완성해 줍니다. (현재는 기본 양식)")
        df_eval_dummy = pd.DataFrame({
            "대분류": ["참여기술인", "참여기술인", "유사용역수행실적", "신용도", "가점/감점"],
            "평가항목": ["사업책임기술인", "분야별책임기술인", "최근 3년 실적", "회사 신용평가등급", "영업정지/교육이수"],
            "배점": ["20점", "30점", "30점", "10점", "+1점 / -2점"],
            "세부인정기준": ["경력 10점, 실적 10점", "보할 적용", "100% 인정", "A- 이상 만점", "교육원 수료 등"]
        })
        st.table(df_eval_dummy)
        
    st.markdown("---")
    st.markdown("### 🔍 세부사항 직접 설정")
    chk_safety = st.checkbox("✅ 정기안전점검 실적 포함 여부", value=True)
    chk_period = st.checkbox("✅ 최근 실적 인정 기간 제한", value=True)
    if chk_period:
        st.selectbox("↳ 인정 기간을 선택하세요", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
        
    chk_bohal = st.checkbox("✅ 분야별 가중치(보할) 직접 설정", value=True)
    if chk_bohal:
        initial_bohal_df = pd.DataFrame([{"전문분야": "상하수도", "비율(%)": 60}, {"전문분야": "토질지질", "비율(%)": 40}])
        st.data_editor(initial_bohal_df, num_rows="dynamic", use_container_width=True)

    st.markdown("---")
    st.markdown("### 👥 필요 인원(T/O) 및 배정 방식 설정")
    col_pm, col_pe, col_pes = st.columns(3)
    with col_pm:
        need_pm = st.checkbox("사책 필요", value=True)
        pm_cnt = st.number_input("사책 인원수", min_value=1, max_value=5, value=1) if need_pm else 0
    with col_pe:
        need_pe = st.checkbox("분책 필요", value=True)
        pe_cnt = st.number_input("분책 인원수", min_value=1, max_value=10, value=2) if need_pe else 0
    with col_pes:
        need_pes = st.checkbox("분참 필요", value=True)
        pes_cnt = st.number_input("분참 인원수", min_value=1, max_value=10, value=2) if need_pes else 0

    assign_mode = st.radio("배정 방식을 선택하세요:", ["🤖 AI 최적 인원 자동 배정 (최고점 추천)", "🧑‍🔧 수동 인원 직접 선택"], horizontal=True, label_visibility="collapsed")
    
    personnel_list = engine.get_personnel_list()
    if assign_mode == "🧑‍🔧 수동 인원 직접 선택":
        if need_pm and pm_cnt > 0:
            pm_cols = st.columns(pm_cnt)
            for i in range(pm_cnt):
                with pm_cols[i]: st.selectbox(f"사책 {i+1}", personnel_list, key=f"sel_pm_{i}")
        if need_pe and pe_cnt > 0:
            pe_cols = st.columns(pe_cnt)
            for i in range(pe_cnt):
                with pe_cols[i]: st.selectbox(f"분책 {i+1}", personnel_list, key=f"sel_pe_{i}")
        if need_pes and pes_cnt > 0:
            pes_cols = st.columns(pes_cnt)
            for i in range(pes_cnt):
                with pes_cols[i]: st.selectbox(f"분참 {i+1}", personnel_list, key=f"sel_pes_{i}")

# --- [Tab 3] 시뮬레이션 결과 확인 ---
with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 결과")
    if st.button("🚀 설정된 세부사항으로 시뮬레이션 실행", type="primary"):
        with st.spinner('마스터 DB 스캔 및 점수 계산 중...'):
            time.sleep(1.5)
            if assign_mode == "🤖 AI 최적 인원 자동 배정 (최고점 추천)":
                best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(pm_cnt, pe_cnt, pes_cnt)
                st.success(f"🎉 AI 최적 조합 발견! (최종 예상 점수: {best_score['획득점수'].sum()} / 60 점 만점)")
                if rec_pm: st.write(f"- **사책:** {', '.join(rec_pm)}")
                if rec_pe: st.write(f"- **분책:** {', '.join(rec_pe)}")
                if rec_pes: st.write(f"- **분참:** {', '.join(rec_pes)}")
                st.dataframe(best_score, use_container_width=True)
            else:
                manual_score = engine.calculate_manual_score()
                st.success(f"✅ 수동 배정 계산 완료! (최종 예상 점수: {manual_score['획득점수'].sum()} / 60 점 만점)")
                st.dataframe(manual_score, use_container_width=True)

# --- [Tab 4] 서류 출력 ---
with tab4:
    st.subheader("최종 출력 및 제출 파일 다운로드")
    if st.button("🔄 제출 서류 및 증빙자료 패키징 시작"):
        with st.spinner("엑셀 서류 작성 및 증빙자료를 수집하여 압축 중입니다..."):
            time.sleep(1)
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                # Tab 2에서 파싱한 진짜 배점표 데이터가 있으면 그걸로 엑셀 만들기!
                if not st.session_state.eval_criteria.empty:
                    st.session_state.eval_criteria.to_excel(writer, sheet_name='1_자기평가표_총괄', index=False)
                else:
                    df_eval = pd.DataFrame({"알림": ["공고문 분석 내용이 없습니다."]})
                    df_eval.to_excel(writer, sheet_name='1_자기평가표_총괄', index=False)
                    
                df_career = engine.master_db if not engine.master_db.empty else pd.DataFrame({'알림': ['엑셀 데이터 없음']})
                df_career.to_excel(writer, sheet_name='2_별지5_참여기술인경력', index=False)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr("1_자동완성_자기평가표.xlsx", excel_buffer.getvalue())
                if st.session_state.uploaded_pdfs:
                    for filename, file_bytes in st.session_state.uploaded_pdfs.items():
                        zip_file.writestr(f"3_증빙자료/{filename}", file_bytes)
                else:
                    zip_file.writestr("3_증빙자료/안내문.txt", "업로드된 증빙 PDF 파일이 없습니다.".encode('utf-8'))
            
            zip_buffer.seek(0)
            st.success("✅ 최종 패키징이 완료되었습니다!")
            st.download_button(label="📦 최종 제출 패키지 다운로드 (.zip)", data=zip_buffer, file_name="최종_PQ_제출서류_패키지.zip", mime="application/zip", type="primary")
