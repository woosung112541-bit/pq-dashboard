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
if 'dream_team' not in st.session_state: st.session_state.dream_team = []
if 'notice_text' not in st.session_state: st.session_state.notice_text = ""
# 💡 [신규] 최종 상세 점수표를 저장할 메모리
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
        
    except Exception as e:
        st.error(f"마스터 DB 다운로드/읽기 오류: {e}")
        load_master_db_from_drive.clear()
        return pd.DataFrame()

def get_all_pdfs_recursively(drive_service, folder_id, depth=0):
    if depth > 5: return [] 
    pdfs = []
    query = f"'{folder_id}' in parents and trashed=false"
    page_token = None
    while True:
        try:
            response = drive_service.files().list(q=query, fields='nextPageToken, files(id, name, mimeType)', pageToken=page_token).execute()
            for file in response.get('files', []):
                if file['mimeType'] == 'application/pdf':
                    pdfs.append(file)
                elif file['mimeType'] == 'application/vnd.google-apps.folder':
                    pdfs.extend(get_all_pdfs_recursively(drive_service, file['id'], depth + 1))
            page_token = response.get('nextPageToken', None)
            if not page_token: break
        except Exception: break
    return pdfs

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
        q_arch = "name='[증빙자료_아카이브]' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_arch = drive_service.files().list(q=q_arch, fields="files(id)").execute()
        if not res_arch.get('files'): return {}
        
        archive_id = res_arch.get('files')[0]['id']
        q_sub = f"'{archive_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        res_sub = drive_service.files().list(q=q_sub, fields="files(id, name)").execute()
        
        archive_status = {}
        for folder in res_sub.get('files', []):
            all_pdfs = get_all_pdfs_recursively(drive_service, folder['id'])
            archive_status[folder['name']] = [f['name'] for f in all_pdfs]
        return archive_status
    except Exception: return {"시스템 스캔 오류 (새로고침을 눌러주세요)": []}

def get_ai_model():
    return genai.GenerativeModel('gemini-3.6-flash')

# ==========================================
# 🧠 [Backend Engine] AI 다이렉트 100점 만점 시뮬레이션 엔진
# ==========================================
class PQScoringEngine:
    def run_ai_dreamteam_optimizer(self, pm_cnt, pe_cnt, pes_cnt, notice_text):
        master_db = load_master_db_from_drive()
        if master_db.empty:
            st.error("❌ 구글 드라이브에서 '마스터'라는 이름이 포함된 엑셀 파일을 찾을 수 없습니다.")
            return pd.DataFrame(), [], [], []
            
        db_csv = master_db.to_csv(index=False)
        
        # 💡 [핵심] 최종 점수표(100점 만점) 양식에 맞춘 프롬프트 업그레이드
        prompt = f"""
        당신은 건설엔지니어링 PQ(사업수행능력평가) 최고 심사위원입니다.
        아래 [공고문 전체 텍스트]와 [엔지니어 실적 Master DB]를 분석하여, 100점 만점 기준의 최종 PQ 점수 산출표와 최적의 '드림팀' 명단을 선발하세요. (단, 가격 점수는 제외하고 서류/실적 점수만 산출합니다.)

        [평가 핵심 지침]
        1. 기간 필터링: 공고 기간(예: 3년)을 2026년 기준으로 정확히 환산하여 유효 실적만 필터링하세요.
        2. 특이 기준 적용: 특정 법령(시안법 등), 특정 분야 가점 등 공고문 내의 모든 세부 규칙을 빠짐없이 DB 실적 점수에 반영하세요.
        3. 전체 항목 평가: 공고문에 명시된 '참여기술인', '유사용역수행실적', '신용도(재정상태/업무정지)', '기술개발및투자실적', '업무중첩도', '가점', '감점' 등 모든 항목의 배점을 파악하고, 이에 맞춘 획득 점수와 구체적인 '점수계산근거'를 작성하세요.
        4. 요구 인원: 사업책임(PM) {pm_cnt}명, 분야책임(PE) {pe_cnt}명, 분야참여(PES) {pes_cnt}명 선발.

        [공고문 전체 텍스트]
        {notice_text[:5000]} 

        [엔지니어 실적 Master DB (CSV 형식)]
        {db_csv}

        분석 결과를 오직 아래 JSON 포맷으로만 반환하세요.
        {{
            "pq_score_table": [
                {{"대분류": "참여기술인", "평가항목": "사업책임기술자", "배점": 20, "획득점수": 20.0, "점수계산근거": "[윤석순] 인정일수 1,200일 이상 (만점)"}},
                {{"대분류": "신용도", "평가항목": "점검진단 실시결과", "배점": 4, "획득점수": 4.0, "점수계산근거": "불량 0건, 매우불량 0건"}},
                {{"대분류": "업무중첩도", "평가항목": "책임기술자", "배점": 6, "획득점수": 6.0, "점수계산근거": "[윤석순] 중복 잔여과업기간 0개월"}}
                // ... 공고문에 명시된 모든 항목에 대해 위 형식으로 작성하여 총점 100점이 되도록 구성할 것 (가/감점 포함)
            ],
            "pm": ["이름1"],
            "pe": ["이름2", "이름3"],
            "pes": ["이름4", "이름5"]
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
st.caption("※ 실무 완벽 대응: 공고문 특이사항 추출 + 100점 만점 상세 산출표 생성기")

tab1, tab2, tab3, tab4 = st.tabs(["📥 1. 마스터 DB 관리", "⚙️ 2. 공고문 설정", "📊 3. 책임기술자 시뮬레이션", "🖨️ 4. 서류 출력 및 패키징"])

# --- [Tab 1] 마스터 DB 및 드라이브 스캔 ---
with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Zone A: 공고문/지침서 분석")
        notice_files = st.file_uploader("공고문 파일(PDF) 업로드", type=['pdf'], accept_multiple_files=True, key="zone_a")
        if notice_files and api_key and st.button("🧠 공고문 AI 분석 및 평가기준 구성", type="primary"):
            with st.spinner("AI가 공고문을 정독하며 숨은 특이사항을 찾아내고 있습니다..."):
                try:
                    notice_text = ""
                    for file in notice_files:
                        pdf = PyPDF2.PdfReader(file)
                        # 평가 기준이 뒷단에 있을 수 있으므로 더 넉넉하게 페이지를 읽어옵니다.
                        for page in pdf.pages[:15]: notice_text += page.extract_text() or ""
                    
                    st.session_state.notice_text = notice_text
                    
                    prompt = f"""
                    건설엔지니어링 PQ 공고문 분석 후 순수 JSON만 반환하세요.
                    1. eval_criteria: 배점표 배열 (대분류, 평가항목, 배점, 세부인정기준 등)
                    2. settings: {{ 
                        has_safety(bool), period(str), bohal(list), pm_cnt(int), pe_cnt(int), pes_cnt(int),
                        extra_settings: {{ "항목명": "내용" }}
                    }}
                    * 중요: extra_settings에는 공고문에 명시된 독특한 가점/감점 기준, 지역 가점, 특정 법령 우대, 신용도 배점 기준 등 모든 세부/특이사항을 자유롭게 추출하세요.
                    공고문: {notice_text}
                    """
                    response = get_ai_model().generate_content(prompt)
                    parsed_json = json.loads(response.text.strip().removeprefix("```json").removesuffix("```").strip())
                    
                    st.session_state.eval_criteria = pd.DataFrame(parsed_json.get("eval_criteria", []))
                    st.session_state.auto_settings = parsed_json.get("settings", st.session_state.auto_settings)
                    st.success("✅ 공고문 분석 성공! 숨겨진 세부사항까지 [Tab 2]에 모두 세팅되었습니다.")
                except Exception as e:
                    st.error(f"공고문 분석 실패: {e}")
                        
    with col2:
        st.subheader("Zone B: 구글 드라이브 실적 아카이브 현황")
        st.info("💡 실적증명서 및 수료증 PDF는 구글 드라이브 `[증빙자료_아카이브]` 폴더에 자유롭게 업로드하시면 됩니다.")
        if st.button("🔍 구글 드라이브 아카이브 현황 새로고침"):
            scan_drive_archive_cached.clear()
            load_master_db_from_drive.clear()
            st.rerun()

        with st.spinner("구글 드라이브 딥-스캔 중..."):
            archive_data = scan_drive_archive_cached()
            if archive_data:
                st.success(f"📂 스캔 완료!")
                for name, pdfs in archive_data.items():
                    with st.expander(f"📁 **{name}** (총 {len(pdfs)}개 서류 보관 중)"):
                        if pdfs:
                            for pdf in pdfs: st.write(f"- 📄 `{pdf}`")
                        else:
                            st.caption("해당 폴더 내부에는 어떠한 PDF도 존재하지 않습니다.")
            else:
                st.warning("구글 드라이브 스캔 실패 또는 폴더가 비어 있습니다.")

# --- [Tab 2] 공고문 세부사항 설정 ---
with tab2:
    st.markdown("### 📊 공고문 AI 분석 결과 및 세부 설정")
    col_title, col_toggle = st.columns([7, 3])
    with col_toggle: manual_override = st.toggle("⚙️ 세부사항 수동 설정", value=False)

    if not st.session_state.eval_criteria.empty: st.table(st.session_state.eval_criteria)
    else: st.info("💡 [Tab 1]에서 공고문을 업로드하시면 평가 배점표가 자동으로 구성됩니다.")
    
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
        with col_b:
            st.write("- **보할 인정 비율:**")
            st.table(pd.DataFrame(bohal_data))
            
        final_pm_cnt, final_pe_cnt, final_pes_cnt = s_settings.get('pm_cnt', 1), s_settings.get('pe_cnt', 0), s_settings.get('pes_cnt', 0)
        
        st.markdown("##### 📌 공고문 특이/세부사항 (AI 자동 추출)")
        if extra_settings:
            for k, v in extra_settings.items():
                st.info(f"**{k}** : {v}")
        else:
            st.caption("별도의 특이사항이 발견되지 않았습니다.")
            
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
    st.markdown("### 🏆 최종 시뮬레이션 및 평가 점수 산출")
    if st.button("🚀 마스터 DB 딥러닝 시뮬레이션 실행 (최종 점수표 생성)", type="primary"):
        if not st.session_state.notice_text:
            st.warning("⚠️ [Tab 1]에서 공고문을 먼저 업로드해야 해당 기준에 맞는 연산이 가능합니다!")
        else:
            with st.spinner('AI가 마스터 DB와 공고문 기준을 매칭하여 100점 만점 기준의 최종 점수 산출표를 작성 중입니다... (약 15초 소요)'):
                best_score_df, rec_pm, rec_pe, rec_pes = engine.run_ai_dreamteam_optimizer(
                    final_pm_cnt, final_pe_cnt, final_pes_cnt, st.session_state.notice_text
                )
                
                if not best_score_df.empty:
                    st.success("🎉 AI 최종 점수 산출 완료!")
                    
                    # 💡 총점 계산 로직 추가
                    total_allocated = best_score_df['배점'].sum()
                    total_earned = best_score_df['획득점수'].sum()
                    
                    col1, col2 = st.columns(2)
                    col1.metric("총 배점 (가격 제외)", f"{total_allocated:g}점")
                    col2.metric("최종 획득 점수", f"{total_earned:g}점")
                    
                    st.dataframe(best_score_df, use_container_width=True)
                    st.session_state.final_pq_score_table = best_score_df
                    
                    st.session_state.dream_team = rec_pm + rec_pe + rec_pes
                    st.info(f"👉 **선발된 드림팀 명단:** {', '.join(st.session_state.dream_team)}")

# --- [Tab 4] 서류 출력 및 패키징 ---
with tab4:
    st.markdown("### 🖨️ 서류 출력 및 자동 패키징")
    if not st.session_state.dream_team:
        st.warning("⚠️ 먼저 [Tab 3]에서 시뮬레이션을 실행하여 기술자 드림팀을 선발해 주세요.")
    else:
        st.success(f"✅ 현재 선발된 기술자 명단: **{', '.join(st.session_state.dream_team)}**")
        if st.button("🔄 구글 드라이브에서 서류 수집 및 ZIP 패키징 시작", type="primary"):
            with st.spinner("드라이브 전역을 뒤져 해당 인원의 모든 증빙 서류를 추출 중입니다..."):
                try:
                    drive_service = authenticate_google_drive()
                    archive_query = "name='[증빙자료_아카이브]' and mimeType='application/vnd.google-apps.folder' and trashed=false"
                    archive_res = drive_service.files().list(q=archive_query, fields="files(id)").execute()
                    
                    if not archive_res.get('files'):
                        st.error("구글 드라이브에 `[증빙자료_아카이브]` 폴더가 존재하지 않습니다.")
                    else:
                        archive_id = archive_res['files'][0]['id']
                        zip_buffer = io.BytesIO()
                        found_files_count = 0
                        folder_map = get_all_subfolders_map(drive_service, archive_id)
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                            # 💡 생성된 점수표도 엑셀로 저장하여 패키지에 포함!
                            if not st.session_state.final_pq_score_table.empty:
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                    st.session_state.final_pq_score_table.to_excel(writer, index=False, sheet_name='최종점수표')
                                z.writestr("0_최종_PQ평가_산출표.xlsx", excel_buffer.getvalue())
                                
                            for person_name in st.session_state.dream_team:
                                if person_name in folder_map:
                                    person_folder_id = folder_map[person_name]
                                    person_pdfs = get_all_pdfs_recursively(drive_service, person_folder_id)
                                    
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
                                    z.writestr(f"{person_name}/안내_폴더없음.txt", f"구글 드라이브에 '{person_name}' 폴더가 어디에도 없습니다.".encode('utf-8'))
                        
                        zip_buffer.seek(0)
                        if found_files_count > 0:
                            st.success(f"🎉 성공! 총 {found_files_count}개의 증빙 PDF와 최종 점수 산출표를 찾아 압축했습니다.")
                            st.download_button("📦 최종 제출서류 패키지 다운로드 (ZIP)", data=zip_buffer, file_name="최종_PQ제출서류_패키지.zip", mime="application/zip", type="primary")
                        else: st.warning("선발된 인원에 해당하는 PDF 서류를 찾지 못했습니다.")
                except Exception as e: st.error(f"서류 패키징 중 오류 발생: {e}")
