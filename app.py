import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import sqlite3
import re
import os
import base64

# =============================================================================
# 0. 기본 설정 및 아이콘 / 데이터 / 배경 파일 경로
# =============================================================================
ICON_FILE = "icon.png"
DB_FILE = "genealogy.db"
DATA_FILE = "genealogy.xlsx"

# 저장소 업로드 배경 이미지 파일 탐색
BG_PATHS = [
    "background.png",
    "background.jpg",
    "bg.png",
    "bg.jpg",
    "artifacts/file_generation/ttl=63d/output/background.png",
    "artifacts/file_generation/ttl=63d/output/bg.png"
]
bg_file = next((p for p in BG_PATHS if os.path.exists(p)), None)

# 저장소 업로드 분파도(branch diagram) 이미지 파일 탐색
BRANCH_DIAGRAM_PATHS = [
    "branch diagram.png",
    "branch_diagram.png",
    "branch-diagram.png",
    "branchdiagram.png",
    "branch diagram.jpg",
    "branch_diagram.jpg",
    "광주안씨_분파도.png",
    "모바일분파도.png",
    "분파도.png"
]
branch_diagram_file = next((p for p in BRANCH_DIAGRAM_PATHS if os.path.exists(p)), "branch diagram.png")

# 폴백 경로 탐색
if not os.path.exists(ICON_FILE):
    alt_icon = "artifacts/file_generation/ttl=63d/output/icon.png"
    if os.path.exists(alt_icon):
        ICON_FILE = alt_icon

DB_PATHS = [
    "genealogy.db",
    "/tmp/genealogy.db",
    "artifacts/file_generation/ttl=63d/output/genealogy.db"
]
db_file_path = next((p for p in DB_PATHS if os.path.exists(p)), DB_FILE)

if not os.path.exists(DATA_FILE):
    if os.path.exists("완귀공파 족보 검색 시스템.xlsx"):
        DATA_FILE = "완귀공파 족보 검색 시스템.xlsx"
    elif os.path.exists("artifacts/file_generation/ttl=63d/output/genealogy.xlsx"):
        DATA_FILE = "artifacts/file_generation/ttl=63d/output/genealogy.xlsx"
    elif os.path.exists("artifacts/file_generation/ttl=63d/output/완귀공파 족보 통합.xlsx"):
        DATA_FILE = "artifacts/file_generation/ttl=63d/output/완귀공파 족보 통합.xlsx"

# ⭐ st.set_page_config는 반드시 모든 st.* 명령보다 가장 먼저 호출되어야 함
st.set_page_config(
    page_title="광주안씨 완귀공파 종중 족보 검색 시스템",
    page_icon=ICON_FILE if os.path.exists(ICON_FILE) else "📜",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ⭐ 이미지 Base64 인코딩을 @st.cache_resource로 메모리에 영구 캐싱하여 부팅 시 디스크 재읽기 제거
@st.cache_resource(show_spinner=False)
def get_base64_image(image_path):
    if image_path and os.path.exists(image_path):
        try:
            with open(image_path, "rb") as f:
                return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
        except Exception:
            return None
    return None

icon_b64 = get_base64_image(ICON_FILE)
bg_b64 = get_base64_image(bg_file) if bg_file else None

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

def dismiss_mobile_keyboard():
    """
    ⭐ 검색 후 스마트폰·태블릿 가상 자판(키보드)을 자동으로 숨기는 함수:
    iframe(components.html) 없이 제로 크기(0x0) 이미지의 onload 핸들러로 현재 활성화된 input 요소를 blur 처리하여
    모바일 인앱 브라우저에서도 무한 로딩 현상 없이 즉시 자판을 내려줍니다.
    """
    render_html(
        '<img src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7" '
        'style="display:none;width:0;height:0;" '
        'onload="try{if(document.activeElement&&document.activeElement.tagName===\'INPUT\'){document.activeElement.blur();}'
        'var ins=document.querySelectorAll(\'input[type=text]\');for(var i=0;i<ins.length;i++){ins[i].blur();}}catch(e){}" />'
    )


def enable_image_lightbox():
    """
    ⭐ 모바일/PC 공용 대형 딥 줌(Deep Zoom & Pan) 라이트박스 뷰어:
    - Streamlit iframe 격리를 넘어 최상위 브라우저 창(window.parent.document)에 직접 주입
    - 화면 크기에 갇히지 않고 최대 8배(800%)까지 화면 밖으로 시원하게 확대
    - 확대된 상태에서 손가락으로 상하좌우 구석구석 자유롭게 드래그(Pan) 이동
    - 스마트폰 더블 탭(톡톡) 시 2.8배 즉시 확대 및 원래 크기 토글
    - 상단 '1:1 원본' 리셋 버튼 및 '✕' 닫기 버튼 제공
    - 80px 미만 작은 아이콘은 제외하고 본문 모든 사진(분파도, 묘소, 비석 등)에 일괄 자동 적용
    """
    components.html("""
    <script>
    (function() {
        try {
            var doc = window.parent.document;
            if (doc._stDeepZoomInitialized) return;
            doc._stDeepZoomInitialized = true;

            // 1. 전체화면 딥 줌 모달 생성 (화면 밖으로 무제한 확대 지원)
            var modal = doc.createElement('div');
            modal.id = 'stDeepZoomModal';
            modal.style.cssText = 'display:none; position:fixed; z-index:99999999; left:0; top:0; width:100vw; height:100vh; background:rgba(0,0,0,0.95); backdrop-filter:blur(8px); justify-content:center; align-items:center; overflow:hidden; touch-action:none;';

            var closeBtn = doc.createElement('div');
            closeBtn.innerHTML = '&times;';
            closeBtn.style.cssText = 'position:absolute; top:18px; right:20px; color:#ffffff; font-size:38px; font-weight:bold; cursor:pointer; z-index:100000000; width:48px; height:48px; border-radius:50%; background:rgba(0,0,0,0.65); display:flex; align-items:center; justify-content:center; line-height:1; user-select:none; box-shadow:0 2px 10px rgba(0,0,0,0.6);';

            var resetBtn = doc.createElement('div');
            resetBtn.innerText = '1:1 맞춤';
            resetBtn.style.cssText = 'position:absolute; top:20px; left:20px; color:#ffffff; font-size:12.5px; font-weight:600; cursor:pointer; z-index:100000000; padding:8px 14px; border-radius:20px; background:rgba(0,0,0,0.65); border:1px solid rgba(255,255,255,0.3); user-select:none;';

            var imgWrapper = doc.createElement('div');
            imgWrapper.style.cssText = 'width:100%; height:100%; display:flex; justify-content:center; align-items:center; overflow:hidden; pointer-events:none;';

            var img = doc.createElement('img');
            img.id = 'stDeepZoomImg';
            img.style.cssText = 'max-width:92vw; max-height:88vh; object-fit:contain; transform-origin:center center; cursor:grab; user-select:none; -webkit-user-drag:none; pointer-events:auto; will-change:transform;';

            imgWrapper.appendChild(img);
            modal.appendChild(closeBtn);
            modal.appendChild(resetBtn);
            modal.appendChild(imgWrapper);
            doc.body.appendChild(modal);

            // 확대(Scale) 및 이동(Position) 상태
            var scale = 1, posX = 0, posY = 0;
            var startDist = 0, startScale = 1;
            var startX = 0, startY = 0;
            var isDragging = false;
            var lastTap = 0;

            function updateTransform() {
                img.style.transform = 'translate(' + posX + 'px, ' + posY + 'px) scale(' + scale + ')';
            }

            function resetZoom() {
                scale = 1; posX = 0; posY = 0;
                updateTransform();
            }

            function closeModal() {
                modal.style.display = 'none';
                resetZoom();
            }

            closeBtn.onclick = closeModal;
            resetBtn.onclick = resetZoom;
            modal.onclick = function(e) {
                if (e.target === modal || e.target === imgWrapper) closeModal();
            };

            // 더블 탭 확대 (화면 밖 2.8배 확대 또는 원래 크기)
            img.addEventListener('touchend', function(e) {
                var currentTime = new Date().getTime();
                var tapLength = currentTime - lastTap;
                if (tapLength < 300 && tapLength > 0) {
                    e.preventDefault();
                    if (scale > 1.2) {
                        resetZoom();
                    } else {
                        scale = 2.8;
                        updateTransform();
                    }
                }
                lastTap = currentTime;
            });

            // 스마트폰 핀치 줌 (최대 8배까지 화면 밖으로 무제한 확대) & 드래그 이동
            modal.addEventListener('touchstart', function(e) {
                if (e.touches.length === 2) {
                    var dx = e.touches[0].clientX - e.touches[1].clientX;
                    var dy = e.touches[0].clientY - e.touches[1].clientY;
                    startDist = Math.hypot(dx, dy);
                    startScale = scale;
                } else if (e.touches.length === 1) {
                    startX = e.touches[0].clientX - posX;
                    startY = e.touches[0].clientY - posY;
                    isDragging = true;
                }
            }, { passive: false });

            modal.addEventListener('touchmove', function(e) {
                e.preventDefault();
                if (e.touches.length === 2) {
                    var dx = e.touches[0].clientX - e.touches[1].clientX;
                    var dy = e.touches[0].clientY - e.touches[1].clientY;
                    var dist = Math.hypot(dx, dy);
                    if (startDist > 0) {
                        scale = Math.min(Math.max(startScale * (dist / startDist), 0.8), 8.0);
                        updateTransform();
                    }
                } else if (e.touches.length === 1 && isDragging) {
                    posX = e.touches[0].clientX - startX;
                    posY = e.touches[0].clientY - startY;
                    updateTransform();
                }
            }, { passive: false });

            modal.addEventListener('touchend', function(e) {
                isDragging = false;
                if (scale < 1
