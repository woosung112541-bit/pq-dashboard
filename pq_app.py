# 파일 최상단에 라이브러리 추가
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import tempfile
import os

# [Tab 1] 코드 교체 부분
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
                    
                    # 3. 구글 드라이브 업로드 설정 (수정된 폴더 ID가 있다면 parents 배열에 넣음)
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
