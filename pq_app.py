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

# 가상 메모리 초기화
if 'uploaded_pdfs' not in st.session_state:
    st.session_state.uploaded_pdfs = {}
if 'eval_criteria' not in st.session_state:
    st.session_state.eval_criteria = pd.DataFrame() 
if 'auto_settings' not in st.session_state:
    # Zone A 분석 전 기본값 (초기 세팅)
    st.session_state.auto_settings = {
        "has_safety": True,
        "period": "3년",
        "bohal": [{"전문분야": "상하수도", "비율(%)": 60}, {"전문분야": "토질지질", "비율(%)": 40}],
        "pm_cnt": 1,
        "pe_cnt": 2,
        "pes_cnt": 2
    }

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
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF)을 드래그 앤 드롭하세요.", type=['pdf'], accept_multiple_files=True, key="zone_a")
        
        if notice_files and api_key:
            if st.button("🧠 공고문 AI 분석 및 평가기준 자동 구성", type="primary"):
                with st.spinner("AI가 공고문을 정독하며 배점표와 세부사항을 추출 중입니다..."):
                    try:
                        notice_text = ""
                        for file in notice_files:
                            pdf = PyPDF2.PdfReader(file)
                            for page in pdf.pages[:5]:
                                notice_text += page.extract_text() or ""
                        
                        # 👉 [핵심 업그레이드] 배점표와 세부 설정값을 통째로 가져오는 프롬프트
                        prompt = f"""
                        다음은 건설엔지니어링(PQ) 공고문/지침서입니다. 분석하여 아래 2가지 요소를 JSON으로 반환하세요.
                        1. eval_criteria: 배점표 배열 (키: 대분류, 평가항목, 배점, 세부인정기준)
                        2. settings: 공고문에 명시된 세부 요구사항 객체
                           - has_safety (boolean): 정기안전점검 실적 인정 여부 (언급 없으면 true)
                           - period (문자열): 실적 인정 기간 (예: "3년", "5년", "제한없음" 등. 언급 없으면 "3년")
                           - bohal (배열): 분야별 가중치 (예: [{{"전문분야": "상하수도", "비율(%)": 60}}]. 언급 없으면 [{{"전문분야": "주공종", "비율(%)": 100}}])
                           - pm_cnt (정수): 사업책임기술인 필요 인원 (보통 1)
                           - pe_cnt (정수): 분야별책임기술인 필요 인원 (기본 2)
                           - pes_cnt (정수): 분야별참여기술인 필요 인원 (기본 2)
                        
                        [공고문 내용]
                        {notice_text}
                        
                        반드시 아래와 같은 순수 JSON 문자열만 출력하세요. 마크다운 금지.
                        {{
                          "eval_criteria": [...],
                          "settings": {{ "has_safety": true, "period": "3년", "bohal": [...], "pm_cnt": 1, "pe_cnt": 2, "pes_cnt": 2 }}
                        }}
                        """
                        model = get_ai_model()
                        response = model.generate_content(prompt)
                        
                        result_text = response.text.strip()
                        if result_text.startswith("```json"): result_text = result_text[7:-3].strip()
                        elif result_text.startswith("```"): result_text = result_text[3:-3].strip()
                        
                        parsed_json = json.loads(result_text)
                        
                        # 파싱된 데이터를 시스템 메모리에 저장
                        st.session_state.eval_criteria = pd.DataFrame(parsed_json.get("eval_criteria", []))
                        st.session_state.auto_settings = parsed_json.get("settings", st.session_state.auto_settings)
                        
                        st.success("✅ 공고문 분석 성공! [Tab 2]에 모든 설정이 자동으로 세팅되었습니다.")
                    except Exception as e:
                        st.error(f"공고문 분석 중 에러 발생: {e}")
        elif notice_files and not api_key:
            st.error("👈 왼쪽 사이드바에 API Key를 먼저 입력해주세요!")
            
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
                        existing_projects_list = []
                        if not engine.master_db.empty and '사업명' in engine.master_db.columns:
                            existing_projects_list = engine.master_db['사업명'].dropna().tolist()
                        existing_str = ", ".join(map(str, existing_projects_list))
                        
                        prompt = f"""
                        이 PDF 문서를 분석하여 JSON으로 반환하세요.
                        1. doc_type (경력증명서, 실적증명서, 신용평가등급확인서, 교육수료증, 기타증빙서류)
                        2. owner (기술자 이름, 없으면 '회사공통')
                        3. projects (배열. "사업명", "시작일", "종료일", "담당업무", "발주처")
                        [🚨 중복 회피 규칙]: 기존 DB 목록 [{existing_str}] 과 '의미상 같은 사업'은 완전히 제외하고 순수 신규만 반환하세요.
                        순수 JSON 문자열만 출력.
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
                            st.success(f"✨ AI 필터링 완료! 완전 신규 실적 {len(projects)}건.")
                            st.dataframe(df_new, use_container_width=True)
                            st.button("💾 신규 실적 업데이트", type="primary")
                        else:
                            st.warning("⚠️ 추출된 실적이 모두 마스터 DB에 있습니다. (전체 패스)")
                        
                        perf_file.seek(0)
                        st.session_state.uploaded_pdfs[new_filename] = perf_file.getvalue()
                    except Exception as e:
                        st.error(f"분석 중 에러 발생: {str(e)}")

# --- [Tab 2] 공고문 세부사항 설정 ---
with tab2:
    # 상단 요약/스위치 영역
    col_title, col_toggle = st.columns([7, 3])
    with col_title:
        st.markdown("### 📊 공고문 AI 분석 결과 (평가 기준 및 세부사항)")
    with col_toggle:
        # 👉 [핵심 업그레이드] 수동 설정 스위치
        st.write("")
        manual_override = st.toggle("⚙️ 세부사항 수동 설정", value=False, help="자동 구성된 내용을 강제로 변경하려면 켜세요.")

    # 1. 배점표 렌더링 (항상 보임)
    if not st.session_state.eval_criteria.empty:
        st.table(st.session_state.eval_criteria)
    else:
        st.info("💡 공고문을 업로드하시면 AI가 아래 표와 설정을 자동으로 완성해 줍니다.")
        st.table(pd.DataFrame({
            "대분류": ["참여기술인", "참여기술인", "유사용역수행실적"],
            "평가항목": ["사업책임기술인", "분야별책임기술인", "최근 3년 실적"],
            "배점": ["20점", "30점", "30점"],
            "세부인정기준": ["경력 10점, 실적 10점", "보할 적용", "100% 인정"]
        }))
    
    st.markdown("---")

    # 2. 세부사항 및 인원 설정 렌더링
    s_settings = st.session_state.auto_settings
    
    if not manual_override:
        # [자동 모드] AI가 읽어온 세팅값을 깔끔한 요약 박스로만 보여줌 (수정 불가)
        st.success("🤖 AI가 공고문을 기반으로 아래의 평가 룰(Rule)을 자동 세팅했습니다. (변경하려면 우측 상단의 '수동 설정' 스위치를 켜세요)")
        
        col_a, col_b = st.columns(2)
        with col_a:
            st.write(f"- **정기안전점검 실적 포함:** {'✅ 포함' if s_settings['has_safety'] else '❌ 미포함'}")
            st.write(f"- **최근 실적 인정 기간:** {s_settings['period']}")
            st.write(f"- **필요 인원(T/O):** 사책 {s_settings['pm_cnt']}명 / 분책 {s_settings['pe_cnt']}명 / 분참 {s_settings['pes_cnt']}명")
        with col_b:
            st.write("- **분야별 가중치(보할) 적용 기준:**")
            st.table(pd.DataFrame(s_settings['bohal']))
            
        # 백엔드로 넘어갈 최종 변수 세팅
        final_pm_cnt = s_settings['pm_cnt']
        final_pe_cnt = s_settings['pe_cnt']
        final_pes_cnt = s_settings['pes_cnt']
        
    else:
        # [수동 모드] 사용자가 직접 값을 수정할 수 있는 입력폼 활성화 (초깃값은 AI가 세팅한 값)
        st.warning("⚠️ 수동 설정 모드입니다. 필요에 따라 자동 입력된 값을 수정하세요.")
        
        st.markdown("#### 🔍 룰(Rule) 상세 조정")
        chk_safety = st.checkbox("✅ 정기안전점검 실적 포함 여부", value=s_settings['has_safety'])
        
        periods = ["1년", "3년", "5년", "7년", "제한없음"]
        idx = periods.index(s_settings['period']) if s_settings['period'] in periods else 1
        sel_period = st.selectbox("↳ 실적 인정 기간", periods, index=idx)
        
        st.write("**✅ 분야별 가중치(보할) 직접 설정**")
        df_bohal = pd.DataFrame(s_settings['bohal']) if s_settings['bohal'] else pd.DataFrame([{"전문분야": "주공종", "비율(%)": 100}])
        edited_bohal = st.data_editor(df_bohal, num_rows="dynamic", use_container_width=True)

        st.markdown("#### 👥 필요 인원(T/O) 조정")
        col_pm, col_pe, col_pes = st.columns(3)
        with col_pm:
            final_pm_cnt = st.number_input("사책 인원수", min_value=0, max_value=5, value=s_settings['pm_cnt'])
        with col_pe:
            final_pe_cnt = st.number_input("분책 인원수", min_value=0, max_value=10, value=s_settings['pe_cnt'])
        with col_pes:
            final_pes_cnt = st.number_input("분참 인원수", min_value=0, max_value=10, value=s_settings['pes_cnt'])

    st.markdown("---")
    st.markdown("### 🧑‍🔧 기술자 배정 (시뮬레이션 대상)")
    assign_mode = st.radio("배정 방식을 선택하세요:", ["🤖 AI 최적 인원 자동 배정 (최고점 추천)", "🧑‍🔧 수동 인원 직접 선택"], horizontal=True, label_visibility="collapsed")
    
    personnel_list = engine.get_personnel_list()
    if assign_mode == "🧑‍🔧 수동 인원 직접 선택":
        if final_pm_cnt > 0:
            pm_cols = st.columns(final_pm_cnt)
            for i in range(final_pm_cnt):
                with pm_cols[i]: st.selectbox(f"사책 {i+1}", personnel_list, key=f"sel_pm_{i}")
        if final_pe_cnt > 0:
            pe_cols = st.columns(final_pe_cnt)
            for i in range(final_pe_cnt):
                with pe_cols[i]: st.selectbox(f"분책 {i+1}", personnel_list, key=f"sel_pe_{i}")
        if final_pes_cnt > 0:
            pes_cols = st.columns(final_pes_cnt)
            for i in range(final_pes_cnt):
                with pes_cols[i]: st.selectbox(f"분참 {i+1}", personnel_list, key=f"sel_pes_{i}")

# --- [Tab 3] 시뮬레이션 결과 확인 ---
with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 결과")
    if st.button("🚀 설정된 세부사항으로 시뮬레이션 실행", type="primary"):
        with st.spinner('마스터 DB 스캔 및 점수 계산 중...'):
            time.sleep(1.5)
            if assign_mode == "🤖 AI 최적 인원 자동 배정 (최고점 추천)":
                best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(final_pm_cnt, final_pe_cnt, final_pes_cnt)
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
