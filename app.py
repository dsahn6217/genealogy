import streamlit as st
import pandas as pd
import openpyxl
import re
import os
import base64

# =============================================================================
# 0. 기본 설정 및 아이콘 / 데이터 파일 경로
# =============================================================================
ICON_FILE = "icon.png"
DATA_FILE = "genealogy.xlsx"

# 폴백 경로 탐색
if not os.path.exists(ICON_FILE):
    alt_icon = "artifacts/file_generation/ttl=63d/output/icon.png"
    if os.path.exists(alt_icon):
        ICON_FILE = alt_icon

if not os.path.exists(DATA_FILE):
    alt_data = "genealogy.xlsx"
    if os.path.exists(alt_data):
        DATA_FILE = alt_data
    elif os.path.exists("artifacts/file_generation/ttl=63d/output/genealogy.xlsx"):
        DATA_FILE = "artifacts/file_generation/ttl=63d/output/genealogy.xlsx"

def get_base64_image(image_path):
    if image_path and os.path.exists(image_path):
        try:
            with open(image_path, "rb") as f:
                return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
        except Exception:
            return None
    return None

icon_b64 = get_base64_image(ICON_FILE)

def render_html(html_content):
    """
    마크다운 파서가 4칸 이상의 들여쓰기 공백을 코드 블록(<pre><code>)으로 오인하는 현상을
    원천 방지하기 위해 모든 줄의 앞뒤 공백을 제거하고 순수 HTML로 안전하게 렌더링합니다.
    """
    if not html_content:
        return
    clean_lines = [line.strip() for line in str(html_content).splitlines() if line.strip()]
    clean_html = chr(10).join(clean_lines)
    st.markdown(clean_html, unsafe_allow_html=True)

# =============================================================================
# 1. 페이지 설정 및 전통 한옥/서책 스타일 CSS (제3장 디자인 & 제4장 폰트/크기)
# =============================================================================
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=Nanum+Myeongjo:wght@400;700&display=swap');

html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', 'Malgun Gothic', sans-serif;
}

.stApp {
    background-color: #faf8f5;
    background-image: radial-gradient(#e5dec9 0.75px, transparent 0.75px);
    background-size: 24px 24px;
}

.title-container {
    text-align: center;
    padding: 5px 5px 5px 5px;
    margin-bottom: 2px;
    border-bottom: 2px solid #8c2d19;
    background: #ffffff;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.04);
}
.title-line-1 {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 16pt;
    font-weight: 700;
    color: #1e3c72;
    margin-bottom: -10px;
    letter-spacing: -0.5px;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 12px;
}
.title-line-2 {
    font-family: 'HCR Dotum', '함초롱돋움', 'Noto Sans KR', sans-serif;
    font-size: 14pt;
    font-weight: 600;
    color: #1a202c;
    margin-bottom: 10px;
    letter-spacing: -0.5px;
}
.title-reference-text {
    font-size: 12pt;
    color: #4a5568;
    font-weight: 600;
    margin-top: 2px;
    background: #f7fafc;
    display: inline-block;
    padding: 4px 18px;
    border-radius: 20px;
    border: 1px solid #e2e8f0;
}

.wangwi-intro-box {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-left: 5px solid #1e3c72;
    border-radius: 8px;
    padding: 16px 20px;
    margin-bottom: 22px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    line-height: 1.5;
}
.wangwi-intro-title {
    font-size: 16pt;
    font-weight: 700;
    color: #1e3c72;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.wangwi-intro-item {
    font-size: 11pt;
    color: #2d3748;
    margin-bottom: 4px;
}
.wangwi-intro-label {
    font-weight: 700;
    color: #1a202c;
    display: inline-block;
    min-width: 65px;
}

.search-guide-right-text {
    font-size: 10pt;
    color: #0369a1;
    background-color: #f0f9ff;
    padding: 6px 12px;
    border-radius: 6px;
    border: 1px solid #bae6fd;
    line-height: 1.45;
}

.result-card {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-top: 4px solid #1e3c72;
    border-radius: 10px;
    padding: 22px;
    margin-top: 20px;
    margin-bottom: 25px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.05);
}
.result-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid #edf2f7;
    padding-bottom: 12px;
    margin-bottom: 15px;
}
.result-name-big {
    font-size: 20pt;
    font-weight: 800;
    color: #1e3c72;
    display: flex;
    align-items: center;
    gap
