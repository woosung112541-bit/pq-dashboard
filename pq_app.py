import streamlit as st
import pandas as pd
import io
import json
import tempfile
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import time

# ==========================================
# 🧠 [Backend Engine] PQ 점수 계산 및 AI 추천 엔진
# ==========================================
class PQScoringEngine:
    def __init__(self):
        # [추후 개발] 구글 드라이브에서 '★책임기술자 마스터 파일.xlsx'를 읽어오는 로직이 들어갈 자리
        # 현재는 테스트를 위해 가상의 사내 DB를 생성합니다.
        self.master_db = pd.DataFrame({
            "이름": ["윤석순", "황흥만", "김진규", "김대리", "이사원", "최부장", "박차장"],
            "직급": ["전무", "상무", "이사", "대리", "사원", "부장", "차장"],
            "전문분야": ["상하수도", "토질지질", "토질지질", "상하수도", "구조", "상하수도", "토목"],
            "경력점수": [17.0, 16.5, 14.0, 8.0, 4.0, 15.0, 12.0],
            "실적점수": [10.0, 9.0, 6.0, 3.0, 1.0, 8.0, 5.0],
            "업무중첩(건)": [0, 1, 0, 0, 0, 2, 0]
        })

    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt):
        """AI가 마스터 DB를 뒤져서 중첩 감점이 없고 점수가 제일 높은 사람들을 뽑아냅니다."""
        # 시뮬레이션을 위해 경력/실적 점수가 높은 순으로 정렬
        sorted_db = self.master_db.sort_values(by=["경력점수", "실적점수"], ascending=[False, False])
        
        # 임시 배정 로직 (T/O 만큼 위에서부터 뽑음)
        pm_list = sorted_db.iloc[0:pm_cnt]['이름'].tolist() if pm_cnt > 0 else []
        pe_list = sorted_db.iloc[pm_cnt : pm_cnt+pe_cnt]['이름'].tolist() if pe_cnt > 0 else []
        pes_list = sorted_db.iloc[pm_cnt+pe_cnt : pm_cnt+pe_cnt+pes_cnt]['이름'].tolist() if pes_cnt > 0 else []
        
        # 최고점 조합 결과 생성
        best_score_df = pd.DataFrame({
            "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
            "배점": [30, 20, 10, -5],
            "획득점수": [30.0, 20.0, 10.0, 0.0],
            "비고": ["AI 드림팀 최적화 (만점)", "중첩도 0건 필터링", "A+ 등급 적용", "해당없음"]
        })
        return best_score_df, pm_list, pe_list, pes_list

    def calculate_manual_score(self):
        """수동으로 선택한 인원의 점수를 DB에서 찾아서 계산합니다."""
        # 더미 계산 결과
        return pd.DataFrame({
            "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
            "배점": [30, 20, 10, -5],
            "획득점수": [27.6, 18.5, 9.8, -1.0],
            "비고": ["일부 인원 경력 미달", "진행중 1건 감점", "BB- 등급", "교육 미이수 감점"]
        })

# 엔진 가동 준비
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

# ---------------------------------------------------------
# [Tab 1] 데이터 입력 및 마스터 DB 관리 (구동 확인 완료)
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# [Tab 2] 공고 룰(Rule) 셋업 (UI/UX 최적화 완료)
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# [Tab 3] 책임기술자 선택 및 PQ 시뮬레이션 (엔진 연동 완료)
# ---------------------------------------------------------
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
        
        # [엔진 가동 버튼 1]
        if st.button("🚀 AI 최적 드림팀 찾기 (시뮬레이션 시작)", type="primary"):
            with st.spinner('마스터 DB 스캔 및 수만 가지 조합 시뮬레이션 중...'):
                time.sleep(1.5) # 연산하는 척 딜레이
                # 엔진 호출!
                best_score, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(pm_cnt, pe_cnt, pes_cnt)
                
                st.success(f"🎉 최적의 조합을 찾았습니다! (최종 예상 점수: {best_score['획득점수'].sum()} / 60 점 만점)")
                st.write("**[AI 추천 드림팀 명단]**")
                if rec_pm: st.write(f"- **사책:** {', '.join(rec_pm)}")
                if rec_pe: st.write(f"- **분책:** {', '.join(rec_pe)}")
                if rec_pes: st.write(f"- **분참:** {', '.join(rec_pes)}")
                
                st.write("**[예상 점수표]**")
                st.dataframe(best_score, use_container_width=True)
                
    else:
        st.markdown("##### 👥 기술자 직접 선택")
        personnel_list = ["(선택)"] + engine.master_db['이름'].tolist()
        
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
                    
        # [엔진 가동 버튼 2]
        if st.button("📊 선택된 인원으로 점수 계산하기", type="primary"):
            with st.spinner('선택 인원 마스터 DB 스캔 및 점수 계산 중...'):
                time.sleep(1)
                # 엔진 호출!
                manual_score = engine.calculate_manual_score()
                
                st.write(f"### 🏆 예상 가채점 결과: {manual_score['획득점수'].sum()} 점 / 60 점 만점")
                st.dataframe(manual_score, use_container_width=True)

# ---------------------------------------------------------
# [Tab 4] 서류 출력 및 증빙 자동 패키징 (유지)
# ---------------------------------------------------------
import zipfile

# ---------------------------------------------------------
# [Tab 4] 서류 출력 및 증빙 자동 패키징 (최종 ZIP 엔진 탑재)
# ---------------------------------------------------------
with tab4:
    st.subheader("최종 출력 및 제출 파일 다운로드")
    st.info("💡 [AI 배정] 또는 [수동 계산]으로 확정된 기술자 명단을 바탕으로 자기평가표가 작성되며, 필요한 증빙 PDF들만 자동으로 모아 ZIP으로 압축합니다.")
    
    # 패키징 엔진 가동
    if st.button("🔄 제출 서류 및 증빙자료 패키징 시작"):
        with st.spinner("엑셀 서류 작성 및 구글 드라이브 증빙자료를 수집하여 압축 중입니다..."):
            time.sleep(2) # 파일 수집(스캔) 딜레이 시뮬레이션
            
            # 1. 엑셀 파일들을 담을 메모리 버퍼 준비
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                # 시뮬레이션된 평가 결과를 '자기평가표' 시트에 작성
                df_eval = pd.DataFrame({
                    "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
                    "획득점수": [30.0, 20.0, 10.0, 0.0],
                    "비고": ["AI 자동 작성", "중첩없음", "A+ 등급", "해당없음"]
                })
                df_eval.to_excel(writer, sheet_name='1_자기평가표_총괄', index=False)
                
                # 기술자 명단을 '별지5_경력사항' 시트에 작성
                df_career = engine.master_db[['이름', '전문분야', '경력점수', '실적점수']]
                df_career.to_excel(writer, sheet_name='2_별지5_참여기술인경력', index=False)
            
            # 2. ZIP 파일 메모리 버퍼 생성 (서버 하드디스크 미사용)
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                # [A] 완성된 엑셀 파일을 ZIP 안에 넣기
                zip_file.writestr("1_자동완성_자기평가표.xlsx", excel_buffer.getvalue())
                
                # [B] 증빙 PDF 파일들을 ZIP 안의 '3_증빙자료' 폴더에 넣기 (가상 파일로 시뮬레이션)
                dummy_pdf_content = b"%PDF-1.4\n%This is a simulated PDF file for evidence."
                zip_file.writestr("3_증빙자료/윤석순_상하수도_실적증명서.pdf", dummy_pdf_content)
                zip_file.writestr("3_증빙자료/황흥만_토질지질_실적증명서.pdf", dummy_pdf_content)
                zip_file.writestr("3_증빙자료/회사_신용평가등급확인서.pdf", dummy_pdf_content)
            
            zip_buffer.seek(0)
            
            st.success("✅ 최종 패키징이 완료되었습니다! 아래 버튼을 눌러 다운로드하세요.")
            
            # 3. 최종 ZIP 파일 다운로드 버튼 노출
            st.download_button(
                label="📦 최종 제출 패키지 다운로드 (.zip)",
                data=zip_buffer,
                file_name="최종_PQ_제출서류_패키지.zip",
                mime="application/zip",
                type="primary"
            )
