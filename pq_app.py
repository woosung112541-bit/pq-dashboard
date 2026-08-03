import streamlit as st
import pandas as pd
import io
import json
import tempfile
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

st.set_page_config(page_title="PQ 자동화 대시보드", layout="wide")
st.title("🏗️ 건설엔지니어링 PQ 자동화 및 시뮬레이션 대시보드")
st.caption("※ 본 페이지는 로컬 및 클라우드 테스트용 프로토타입입니다.")

# 👇 이 부분이 지워져서 에러가 났던 것입니다! (탭 4개 생성)
tab1, tab2, tab3, tab4 = st.tabs([
    "📥 1. 마스터 DB 관리", 
    "⚙️ 2. 공고 룰(Rule) 셋업", 
    "📊 3. 책임기술자 선택 및 시뮬레이션", 
    "🖨️ 4. 서류 출력 및 증빙 패키징"
])

# ---------------------------------------------------------
# [Tab 1] 데이터 입력 및 마스터 DB 관리 (구글 드라이브 연동 완료)
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
                    # 1. 스트림릿 금고에서 구글 열쇠(JSON) 꺼내기
                    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
                    creds = Credentials.from_service_account_info(
                        creds_dict, scopes=['https://www.googleapis.com/auth/drive']
                    )
                    drive_service = build('drive', 'v3', credentials=creds)
                    
                    # 2. 임시 파일로 저장
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(perf_file.getvalue())
                        tmp_path = tmp.name
                    
                    # 3. 구글 드라이브 업로드 설정
                    file_metadata = {'name': perf_file.name}
                    media = MediaFileUpload(tmp_path, mimetype='application/pdf')
                    
                    # 4. 파일 업로드 실행
                    uploaded_file = drive_service.files().create(
                        body=file_metadata, media_body=media, fields='id'
                    ).execute()
                    
                    os.remove(tmp_path) # 임시 파일 삭제
                    st.success(f"✅ 구글 드라이브 업로드 완료! (파일 ID: {uploaded_file.get('id')})")
                    
                except Exception as e:
                    st.error(f"업로드 중 에러가 발생했습니다: {str(e)}")

# ---------------------------------------------------------
# [Tab 2] 공고 룰(Rule) 셋업
# ---------------------------------------------------------
# ---------------------------------------------------------
# [Tab 2] 공고 룰(Rule) 셋업
# ---------------------------------------------------------
with tab2:
    st.subheader("평가 기준(Rule) 확정")
    rule_option = st.radio(
        "룰(Rule) 세팅 방식을 선택하세요:", 
        ("1. 발주처별 세팅된 룰 불러오기", "2. 세부사항 선택 및 더블체크 (직접 설정)"),
        horizontal=True
    )
    
    if rule_option == "1. 발주처별 세팅된 룰 불러오기":
        st.selectbox("미리 세팅된 래퍼런스 선택", ["한국환경공단 (1억 미만)", "한국수자원공사 (사후 PQ)", "한국농어촌공사"])
    else:
        st.markdown("#### 🔍 AI 파싱 결과 더블체크 (Human-in-the-loop)")
        col_a, col_b = st.columns(2)
        
        with col_a:
            st.checkbox("정기안전점검 실적 포함 여부", value=True)
            # 👉 수정 1: 실적 인정 기간 선택지 세분화
            st.selectbox("최근 실적 인정 기간", ["1년", "3년", "5년", "7년", "제한없음"], index=1)
            
        with col_b:
            # 👉 수정 2: 보할 유무 선택 및 동적 추가 에디터
            st.write("**분야별 보할(가중치) 설정**")
            has_bohal = st.radio("보할(가중치) 적용 여부", ["적용 안 함", "분야별 보할 적용 (직접 설정)"], horizontal=True)
            
            if has_bohal == "분야별 보할 적용 (직접 설정)":
                st.caption("※ 표의 맨 아래 빈칸을 클릭하여 분야명(예: 교량, 터널)을 입력하고 비율을 추가하세요.")
                
                # 초기 설정값 (사용자가 마음대로 지우고 덮어쓸 수 있음)
                initial_bohal_df = pd.DataFrame([
                    {"전문분야": "상하수도", "비율(%)": 60},
                    {"전문분야": "토질지질", "비율(%)": 40}
                ])
                
                # 데이터를 자유롭게 추가/수정/삭제할 수 있는 동적 에디터
                edited_bohal_df = st.data_editor(
                    initial_bohal_df,
                    num_rows="dynamic",
                    use_container_width=True
                )
                
                # 비율 합계 자동 검증 로직
                total_bohal = edited_bohal_df["비율(%)"].sum()
                if total_bohal != 100:
                    st.warning(f"⚠️ 현재 보할 합계: {total_bohal}% (100%로 맞춰주세요)")
                else:
                    st.success(f"✅ 합계 100% 완료")
                    
        st.button("저장 및 룰 확정")

# ---------------------------------------------------------
# [Tab 3] 책임기술자 선택 및 PQ 시뮬레이션
# ---------------------------------------------------------
with tab3:
    st.subheader("평가 대상 기술자 선택")
    col_x, col_y, col_z = st.columns(3)
    
    with col_x:
        st.selectbox("사업책임기술인", ["(선택)", "윤석순", "김대리"])
    with col_y:
        st.selectbox("분야별책임 (상하수도)", ["(공란)", "윤석순", "김대리"])
    with col_z:
        st.selectbox("분야별책임 (토질지질)", ["(공란)", "황흥만", "김진규"])
        
    st.markdown("---")
    if st.button("📊 점수 계산하기 (시뮬레이션 시작)", type="primary"):
        with st.spinner('마스터 DB를 분석하여 점수를 계산 중입니다...'):
            st.write("### 🏆 최종 예상 가채점 결과: 57.4 점 / 60 점 만점")
            df_result = pd.DataFrame({
                "평가항목": ["사업수행능력", "업무중첩도", "신용도", "감점"],
                "배점": [30, 20, 10, -5],
                "획득점수": [27.6, 20.0, 9.8, 0.0],
                "비고": ["윤석순 경력 17점", "진행중 0건", "BB- 등급", "해당없음"]
            })
            st.dataframe(df_result, use_container_width=True)

# ---------------------------------------------------------
# [Tab 4] 서류 출력 및 증빙 자동 패키징
# ---------------------------------------------------------
with tab4:
    st.subheader("최종 출력 및 제출 파일 다운로드")
    st.info("선택된 기술자의 실적을 기반으로 서류가 작성되며, 관련 증빙 PDF를 자동으로 스캔하여 압축합니다.")
    
    # 더미 엑셀 파일 생성 (다운로드 테스트용)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame({'테스트': ['이 파일은 자동 생성된 자기평가표 샘플입니다.']}).to_excel(writer, index=False)
    output.seek(0)
    
    st.download_button(
        label="📦 최종 제출 패키지 다운로드",
        data=output,
        file_name="최종제출패키지_자동완성.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
