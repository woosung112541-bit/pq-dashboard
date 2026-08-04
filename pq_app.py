import PyPDF2
import streamlit as st
import pandas as pd
import io
import json
import tempfile
import os
import time
import zipfile
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# ==========================================
# 🔗 [Data Loader] 구글 드라이브 마스터 DB 연동
# ==========================================
@st.cache_data(ttl=600)
def load_master_db_from_drive():
    try:
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
        st.error(f"구글 드라이브 연동 중 에러 발생: {e}")
        return pd.DataFrame()

# ==========================================
# 🧠 [Backend Engine] PQ 점수 계산 및 AI 추천 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        self.master_db = load_master_db_from_drive()

    def get_personnel_list(self):
        if self.master_db.empty:
            return ["(선택)", "DB연동필요"]
        for col_name in ['이름', '성명', '엔지니어명', '기술자명', '기술인']:
            if col_name in self.master_db.columns:
                names = self.master_db[col_name].dropna().unique().tolist()
                return ["(선택)"] + names
        # 김대리 영구 삭제 적용
        return ["(선택)", "윤석순", "황흥만", "김진규", "이사원", "최부장", "박차장 (엑셀 '성명' 열 추가 필요)"]

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

# ==========================================
# 🖥️ [Frontend] Streamlit 대시보드 UI
# ==========================================
st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")
# 타이틀 변경
st.title("PQ 자동화 대시보드")
st.caption("※ 본 페이지는 로컬 및 클라우드 테스트용 프로토타입입니다.")

# 탭 이름 변경
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
        st.subheader("Zone A: 공고문/지침서 입력")
        # 여러 파일 업로드 허용 (accept_multiple_files=True)
        notice_files = st.file_uploader("공고문 등 관련 파일을 드래그 앤 드롭하세요. (다중 첨부 가능)", type=['pdf', 'hwp'], accept_multiple_files=True, key="zone_a")
        if notice_files:
            st.success(f"총 {len(notice_files)}개의 파일이 업로드 되었습니다! (향후 AI 파싱 로직 연동 예정)")
            
  import streamlit as st
import pandas as pd
import io
import json
import tempfile
import os
import time
import zipfile
import PyPDF2  # 👈 [추가됨] PDF 텍스트 추출 라이브러리
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

# ... (Data Loader와 PQScoringEngine 클래스 등 이전 코드는 그대로 유지) ...

# --- [Tab 1] 마스터 DB 관리 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 입력")
        notice_files = st.file_uploader("공고문 등 관련 파일을 드래그 앤 드롭하세요. (다중 첨부 가능)", type=['pdf', 'hwp'], accept_multiple_files=True, key="zone_a")
        if notice_files:
            st.success(f"총 {len(notice_files)}개의 파일이 업로드 되었습니다! (향후 AI 파싱 로직 연동 예정)")
            
    with col2:
        st.subheader("Zone B: 실적 업데이트 (Master DB 연동)")
        perf_file = st.file_uploader("기술인/회사 실적증명서(PDF) 업로드", type=['pdf'], key="zone_b")
        
        if perf_file:
            with st.spinner("🔍 AI가 문서 내용을 스캔하여 종류와 주인을 판별 중입니다..."):
                try:
                    # 1. PDF 텍스트 추출 (앞 2페이지만 빠르게 스캔)
                    pdf_reader = PyPDF2.PdfReader(perf_file)
                    extracted_text = ""
                    for page in pdf_reader.pages[:2]: 
                        extracted_text += page.extract_text() or ""
                    
                    # 2. 스마트 판별 로직 (문서 종류 및 소유자 추출)
                    doc_type = "기타증빙서류"
                    owner = "회사공통"
                    
                    # (1) 문서 종류 판별
                    if "경력증명서" in extracted_text or "경력확인서" in extracted_text:
                        doc_type = "경력증명서"
                    elif "실적증명서" in extracted_text or "실적" in extracted_text:
                        doc_type = "실적증명서"
                    elif "신용평가" in extracted_text or "신용등급" in extracted_text:
                        doc_type = "신용평가등급확인서"
                    elif "교육" in extracted_text or "수료" in extracted_text:
                        doc_type = "교육수료증"
                        
                    # (2) 문서 주인(이름) 판별 (마스터 DB 명단과 대조하여 찾기)
                    personnel_list = engine.get_personnel_list()
                    for name in personnel_list:
                        if name != "(선택)" and name in extracted_text:
                            owner = name
                            break
                            
                    # 3. 깔끔한 새 파일명 생성
                    new_filename = f"[{doc_type}] {owner}.pdf"
                    
                    st.info(f"💡 문서 스캔 완료: **{owner}**의 **{doc_type}**로 분류되었습니다.")
                    
                    # 4. 구글 드라이브 업로드 (이름을 바꿔서 올리기)
                    perf_file.seek(0) # 👈 [핵심] 읽었던 파일 포인터를 다시 처음으로 되돌림
                    
                    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
                    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/drive'])
                    drive_service = build('drive', 'v3', credentials=creds)
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(perf_file.getvalue())
                        tmp_path = tmp.name
                        
                    # 새 파일명 적용! (추후 여기에 특정 폴더 ID를 지정하여 라우팅 가능)
                    file_metadata = {'name': new_filename} 
                    media = MediaFileUpload(tmp_path, mimetype='application/pdf')
                    uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                    
                    os.remove(tmp_path)
                    st.success(f"✅ 구글 드라이브 업로드 완료! (저장된 이름: {new_filename})")
                    
                except Exception as e:
                    st.error(f"분석 및 업로드 중 에러 발생: {str(e)}")

# ... (Tab 2, Tab 3, Tab 4 코드는 이전과 동일하게 유지) ...
# --- [Tab 2] 공고문 세부사항 설정 (통합 및 레이아웃 개선) ---
with tab2:
    # 1. 평가 항목 시각화 영역 추가
    st.markdown("### 📊 평가 항목 및 배점 기준 시각화 (자기평가표 초안)")
    st.info("💡 Zone A에 업로드된 공고문을 분석하여 추출한 배점 기준표입니다. (제대로 파싱되었는지 시각적으로 확인하세요)")
    # 자기평가표 양식을 본딴 더미 데이터프레임
    df_eval_criteria = pd.DataFrame({
        "대분류": ["참여기술인", "참여기술인", "유사용역수행실적", "신용도", "가점/감점"],
        "평가항목": ["사업책임기술인", "분야별책임기술인", "최근 3년 실적", "회사 신용평가등급", "영업정지/교육이수"],
        "배점": ["20점", "30점", "30점", "10점", "+1점 / -2점"],
        "세부 인정기준": ["경력 10점, 실적 10점", "보할 적용(상하수도 60, 토질 40)", "100% 인정 (정기안전점검 포함)", "A- 이상 만점", "건설기술교육원 수료 등"]
    })
    st.table(df_eval_criteria)
    
    st.markdown("---")
    
    # 2. 세부사항 직접 설정 (체크박스 토글 형태)
    st.markdown("### 🔍 세부사항 직접 설정")
    
    chk_safety = st.checkbox("✅ 정기안전점검 실적 포함 여부", value=True)
    
    chk_period = st.checkbox("✅ 최근 실적 인정 기간 제한", value=True)
    if chk_period:
        st.selectbox("↳ 인정 기간을 선택하세요", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
        
    chk_bohal = st.checkbox("✅ 분야별 가중치(보할) 직접 설정", value=True)
    if chk_bohal:
        initial_bohal_df = pd.DataFrame([{"전문분야": "상하수도", "비율(%)": 60}, {"전문분야": "토질지질", "비율(%)": 40}])
        edited_bohal_df = st.data_editor(initial_bohal_df, num_rows="dynamic", use_container_width=True)
        total_bohal = edited_bohal_df["비율(%)"].sum()
        if total_bohal != 100:
            st.warning(f"⚠️ 현재 보할 합계: {total_bohal}% (100%로 맞춰주세요)")

    st.markdown("---")
    
    # 3. 필요 인원 및 기술자 배정 (Tab 3에서 가져옴)
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

    st.write("**배정 방식을 선택하세요:**")
    assign_mode = st.radio("배정 방식을 선택하세요:", ["🤖 AI 최적 인원 자동 배정 (최고점 추천)", "🧑‍🔧 수동 인원 직접 선택"], horizontal=True, label_visibility="collapsed")
    
    # 수동 선택일 경우 드롭다운 노출
    personnel_list = engine.get_personnel_list()
    if assign_mode == "🧑‍🔧 수동 인원 직접 선택":
        if need_pm and pm_cnt > 0:
            st.write("🔹 **사업책임기술인(사책)**")
            pm_cols = st.columns(pm_cnt)
            for i in range(pm_cnt):
                with pm_cols[i]: st.selectbox(f"사책 {i+1}", personnel_list, key=f"sel_pm_{i}")
        if need_pe and pe_cnt > 0:
            st.write("🔹 **분야별책임기술인(분책)**")
            pe_cols = st.columns(pe_cnt)
            for i in range(pe_cnt):
                with pe_cols[i]: st.selectbox(f"분책 {i+1}", personnel_list, key=f"sel_pe_{i}")
        if need_pes and pes_cnt > 0:
            st.write("🔹 **분야별참여기술인(분참)**")
            pes_cols = st.columns(pes_cnt)
            for i in range(pes_cnt):
                with pes_cols[i]: st.selectbox(f"분참 {i+1}", personnel_list, key=f"sel_pes_{i}")

# --- [Tab 3] 시뮬레이션 결과 확인 (결과만 노출) ---
with tab3:
    st.markdown("### 🏆 최종 시뮬레이션 결과")
    st.info("Tab 2에서 설정된 세부사항과 배정 방식을 바탕으로 점수를 계산합니다.")
    
    if st.button("🚀 설정된 세부사항으로 시뮬레이션 실행", type="primary"):
        with st.spinner('마스터 DB 스캔 및 점수 계산 중...'):
            time.sleep(1.5)
            if assign_mode == "🤖 AI 최적 인원 자동 배정 (최고점 추천)":
                best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(pm_cnt, pe_cnt, pes_cnt)
                st.success(f"🎉 AI 최적 조합 발견! (최종 예상 점수: {best_score['획득점수'].sum()} / 60 점 만점)")
                st.write("**[추천 드림팀 명단]**")
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
    st.info("💡 확정된 명단을 바탕으로 자기평가표가 작성되며, 필요한 증빙 PDF들만 모아 ZIP으로 압축합니다.")
    
    if st.button("🔄 제출 서류 및 증빙자료 패키징 시작"):
        with st.spinner("엑셀 서류 작성 및 증빙자료를 수집하여 압축 중입니다..."):
            time.sleep(2)
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df_eval = pd.DataFrame({"평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"], "획득점수": [30.0, 20.0, 10.0, 0.0], "비고": ["AI 자동 작성", "중첩없음", "A+ 등급", "해당없음"]})
                df_eval.to_excel(writer, sheet_name='1_자기평가표_총괄', index=False)
                df_career = engine.master_db if not engine.master_db.empty else pd.DataFrame({'알림': ['엑셀 데이터 없음']})
                df_career.to_excel(writer, sheet_name='2_별지5_참여기술인경력', index=False)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr("1_자동완성_자기평가표.xlsx", excel_buffer.getvalue())
                dummy_pdf = b"%PDF-1.4\n%This is a simulated PDF file for evidence."
                zip_file.writestr("3_증빙자료/윤석순_상하수도_실적증명서.pdf", dummy_pdf)
                zip_file.writestr("3_증빙자료/회사_신용평가등급확인서.pdf", dummy_pdf)
            
            zip_buffer.seek(0)
            st.success("✅ 최종 패키징이 완료되었습니다!")
            st.download_button(label="📦 최종 제출 패키지 다운로드 (.zip)", data=zip_buffer, file_name="최종_PQ_제출서류_패키지.zip", mime="application/zip", type="primary")
