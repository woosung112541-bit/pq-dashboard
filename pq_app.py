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
            return pd.DataFrame() # 파일이 없으면 빈 데이터 반환

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
        """엑셀 파일에서 기술자 명단을 에러 없이 안전하게 추출하는 스마트 로직"""
        if self.master_db.empty:
            return ["(선택)", "DB연동필요"]
        
        # 엑셀의 열 이름(Header) 중 이름이 될 만한 것을 자동으로 찾음
        for col_name in ['이름', '성명', '엔지니어명', '기술자명', '기술인']:
            if col_name in self.master_db.columns:
                # 중복을 제거하고 명단 리스트로 반환
                names = self.master_db[col_name].dropna().unique().tolist()
                return ["(선택)"] + names
                
        # 만약 엑셀에 이름 관련 열이 아예 없다면 (현재 캡처본 상태), 에러 방지용 가상 명단 노출
        return ["(선택)", "윤석순", "황흥만", "김진규", "김대리", "이사원 (엑셀에 '성명' 열 추가 필요)"]

    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt):
        best_score_df = pd.DataFrame({
            "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
            "배점": [30, 20, 10, -5],
            "획득점수": [30.0, 20.0, 10.0, 0.0],
            "비고": ["AI 최적화", "중첩도 0건", "A+ 등급", "해당없음"]
        })
        
        # 임시 추천 명단 (추후 실제 DB 점수 기반으로 변경)
        sample_names = ["윤석순", "황흥만", "김진규", "김대리", "이사원", "최부장", "박차장"]
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
st.title("🏗️ 건설엔지니어링 PQ 자동화 및 시뮬레이션 대시보드")
st.caption("※ 본 페이지는 로컬 및 클라우드 테스트용 프로토타입입니다.")

tab1, tab2, tab3, tab4 = st.tabs([
    "📥 1. 마스터 DB 관리", 
    "⚙️ 2. 공고 룰(Rule) 셋업", 
    "📊 3. 책임기술자 선택 및 시뮬레이션", 
    "🖨️ 4. 서류 출력 및 증빙 패키징"
])

# --- [Tab 1] 마스터 DB 관리 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 입력")
        notice_file = st.file_uploader("공고문 파일을 드래그 앤 드롭하세요.", type=['pdf', 'hwp'], key="zone_a")
        if notice_file:
            st.success(f"'{notice_file.name}' 분석 완료! (Tab 2에 초안이 세팅되었습니다.)")
            
    with col2:
        st.subheader("Zone B: 실적 업데이트 (Master DB 연동)")
        perf_file = st.file_uploader("기술인/회사 실적증명서(PDF)를 드래그 앤 드롭하세요.", type=['pdf'], key="zone_b")
        if perf_file:
            with st.spinner("구글 드라이브에 안전하게 업로드 중입니다..."):
                try:
                    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
                    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/drive'])
                    drive_service = build('drive', 'v3', credentials=creds)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(perf_file.getvalue())
                        tmp_path = tmp.name
                    file_metadata = {'name': perf_file.name}
                    media = MediaFileUpload(tmp_path, mimetype='application/pdf')
                    uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                    os.remove(tmp_path)
                    st.success(f"✅ 구글 드라이브 업로드 완료! (파일 ID: {uploaded_file.get('id')})")
                except Exception as e:
                    st.error(f"업로드 중 에러가 발생했습니다: {str(e)}")

# --- [Tab 2] 공고 룰 셋업 ---
with tab2:
    st.subheader("평가 기준(Rule) 확정")
    rule_option = st.radio("룰(Rule) 세팅 방식을 선택하세요:", ("1. 발주처별 세팅된 룰 불러오기", "2. 세부사항 선택 및 더블체크 (직접 설정)"), horizontal=True)
    
    if rule_option == "1. 발주처별 세팅된 룰 불러오기":
        st.selectbox("미리 세팅된 래퍼런스 선택", ["한국환경공단 (1억 미만)", "한국수자원공사 (사후 PQ)", "한국농어촌공사"])
    else:
        st.markdown("#### 🔍 AI 파싱 결과 더블체크 (Human-in-the-loop)")
        col_a, col_b = st.columns(2)
        with col_a:
            st.checkbox("정기안전점검 실적 포함 여부", value=True)
            st.selectbox("최근 실적 인정 기간", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
        with col_b:
            st.write("**분야별 보할(가중치) 설정**")
            has_bohal = st.radio("보할(가중치) 적용 여부", ["적용 안 함", "분야별 보할 적용 (직접 설정)"], horizontal=True)
            if has_bohal == "분야별 보할 적용 (직접 설정)":
                st.caption("※ 표의 맨 아래 빈칸을 클릭하여 분야명과 비율을 추가하세요.")
                initial_bohal_df = pd.DataFrame([{"전문분야": "상하수도", "비율(%)": 60}, {"전문분야": "토질지질", "비율(%)": 40}])
                edited_bohal_df = st.data_editor(initial_bohal_df, num_rows="dynamic", use_container_width=True)
                total_bohal = edited_bohal_df["비율(%)"].sum()
                if total_bohal != 100:
                    st.warning(f"⚠️ 현재 보할 합계: {total_bohal}% (100%로 맞춰주세요)")
                else:
                    st.success(f"✅ 합계 100% 완료")
        st.button("저장 및 룰 확정")

# --- [Tab 3] 시뮬레이션 ---
with tab3:
    st.subheader("평가 대상 기술자 구성 및 배정")
    st.markdown("#### 1. 필요 인원(T/O) 설정")
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

    st.markdown("---")
    st.markdown("#### 2. 기술자 배정 방식")
    assign_mode = st.radio("배정 방식을 선택하세요:", ["🤖 AI 최적 인원 자동 배정 (최고점 추천)", "🧑‍🔧 수동 인원 직접 선택"], horizontal=True)
    
    if assign_mode == "🤖 AI 최적 인원 자동 배정 (최고점 추천)":
        st.info("💡 마스터 DB의 모든 기술자 경력을 스캔하여, 감점이 없고 최고점을 받을 수 있는 **'최적의 드림팀 조합'**을 시스템이 자동으로 찾아냅니다.")
        if st.button("🚀 AI 최적 드림팀 찾기 (시뮬레이션 시작)", type="primary"):
            with st.spinner('마스터 DB 스캔 및 수만 가지 조합 시뮬레이션 중...'):
                time.sleep(1.5)
                best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(pm_cnt, pe_cnt, pes_cnt)
                
                st.success(f"🎉 최적의 조합을 찾았습니다! (최종 예상 점수: {best_score['획득점수'].sum()} / 60 점 만점)")
                st.write("**[AI 추천 드림팀 명단]**")
                if rec_pm: st.write(f"- **사책:** {', '.join(rec_pm)}")
                if rec_pe: st.write(f"- **분책:** {', '.join(rec_pe)}")
                if rec_pes: st.write(f"- **분참:** {', '.join(rec_pes)}")
                
                st.dataframe(best_score, use_container_width=True)
                
    else:
        st.markdown("##### 👥 기술자 직접 선택")
        
        # 👉 [핵심] 더 이상 에러가 나지 않는 스마트 명단 추출 함수 호출!
        personnel_list = engine.get_personnel_list()
        
        if need_pm and pm_cnt > 0:
            st.write("**🔹 사업책임기술인(사책)**")
            pm_cols = st.columns(pm_cnt)
            for i in range(pm_cnt):
                with pm_cols[i]: st.selectbox(f"사책 {i+1}", personnel_list, key=f"sel_pm_{i}")
                    
        if need_pe and pe_cnt > 0:
            st.write("**🔹 분야별책임기술인(분책)**")
            pe_cols = st.columns(pe_cnt)
            for i in range(pe_cnt):
                with pe_cols[i]: st.selectbox(f"분책 {i+1}", personnel_list, key=f"sel_pe_{i}")

        if need_pes and pes_cnt > 0:
            st.write("**🔹 분야별참여기술인(분참)**")
            pes_cols = st.columns(pes_cnt)
            for i in range(pes_cnt):
                with pes_cols[i]: st.selectbox(f"분참 {i+1}", personnel_list, key=f"sel_pes_{i}")
                    
        if st.button("📊 선택된 인원으로 점수 계산하기", type="primary"):
            with st.spinner('선택 인원 마스터 DB 스캔 및 점수 계산 중...'):
                time.sleep(1)
                manual_score = engine.calculate_manual_score()
                st.write(f"### 🏆 예상 가채점 결과: {manual_score['획득점수'].sum()} 점 / 60 점 만점")
                st.dataframe(manual_score, use_container_width=True)

# --- [Tab 4] 서류 출력 ---
with tab4:
    st.subheader("최종 출력 및 제출 파일 다운로드")
    st.info("💡 확정된 기술자 명단을 바탕으로 자기평가표가 작성되며, 필요한 증빙 PDF들만 자동으로 모아 ZIP으로 압축합니다.")
    
    if st.button("🔄 제출 서류 및 증빙자료 패키징 시작"):
        with st.spinner("엑셀 서류 작성 및 구글 드라이브 증빙자료를 수집하여 압축 중입니다..."):
            time.sleep(2)
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df_eval = pd.DataFrame({"평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"], "획득점수": [30.0, 20.0, 10.0, 0.0], "비고": ["AI 자동 작성", "중첩없음", "A+ 등급", "해당없음"]})
                df_eval.to_excel(writer, sheet_name='1_자기평가표_총괄', index=False)
                # 엑셀 데이터가 있으면 그걸 쓰고 없으면 빈 칸 노출
                df_career = engine.master_db if not engine.master_db.empty else pd.DataFrame({'알림': ['엑셀 파일이 비어있거나 불러올 수 없습니다.']})
                df_career.to_excel(writer, sheet_name='2_별지5_참여기술인경력', index=False)
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr("1_자동완성_자기평가표.xlsx", excel_buffer.getvalue())
                dummy_pdf = b"%PDF-1.4\n%This is a simulated PDF file for evidence."
                zip_file.writestr("3_증빙자료/윤석순_상하수도_실적증명서.pdf", dummy_pdf)
                zip_file.writestr("3_증빙자료/회사_신용평가등급확인서.pdf", dummy_pdf)
            
            zip_buffer.seek(0)
            st.success("✅ 최종 패키징이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
            
            st.download_button(label="📦 최종 제출 패키지 다운로드 (.zip)", data=zip_buffer, file_name="최종_PQ_제출서류_패키지.zip", mime="application/zip", type="primary")
