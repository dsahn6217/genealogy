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
                if (scale < 1) resetZoom();
            });

            // PC 마우스 드래그 및 휠 줌 (최대 8배)
            var mouseStartX = 0, mouseStartY = 0, isMouseDown = false;
            img.addEventListener('mousedown', function(e) {
                isMouseDown = true;
                mouseStartX = e.clientX - posX;
                mouseStartY = e.clientY - posY;
                img.style.cursor = 'grabbing';
                e.preventDefault();
            });
            doc.addEventListener('mousemove', function(e) {
                if (!isMouseDown) return;
                posX = e.clientX - mouseStartX;
                posY = e.clientY - mouseStartY;
                updateTransform();
            });
            doc.addEventListener('mouseup', function(e) {
                if (isMouseDown) {
                    isMouseDown = false;
                    img.style.cursor = 'grab';
                }
            });
            modal.addEventListener('wheel', function(e) {
                e.preventDefault();
                var delta = e.deltaY < 0 ? 1.25 : 0.8;
                scale = Math.min(Math.max(scale * delta, 0.8), 8.0);
                updateTransform();
            }, { passive: false });

            // Streamlit 부모 문서의 모든 이미지 클릭 감지
            doc.addEventListener('click', function(e) {
                var t = e.target;
                if (t && t.tagName === 'IMG') {
                    if (t.id === 'stDeepZoomImg') return;
                    if (t.classList.contains('no-zoom')) return;
                    if (t.naturalWidth < 80 || t.naturalHeight < 80) return;
                    
                    img.src = t.src;
                    modal.style.display = 'flex';
                    resetZoom();
                }
            }, true);
        } catch(e) {
            console.error("DeepZoom error:", e);
        }
    })();
    </script>
    """, height=0)

def get_query_param(key):
    try:
        if hasattr(st, "query_params") and key in st.query_params:
            val = st.query_params[key]
            if isinstance(val, (list, tuple)):
                return val[0] if val else None
            return str(val)
        elif hasattr(st, "experimental_get_query_params"):
            params = st.experimental_get_query_params()
            if key in params and params[key]:
                return str(params[key][0])
    except Exception:
        pass
    return None

def clear_tree_query_params():
    try:
        if hasattr(st, "query_params"):
            for k in ("tree_gid", "init_gid"):
                if k in st.query_params:
                    del st.query_params[k]
        elif hasattr(st, "experimental_set_query_params"):
            st.experimental_set_query_params()
    except Exception:
        pass

# =============================================================================
# 1. 디자인 스타일 CSS
# =============================================================================
CUSTOM_CSS = """
<style>
html, body, [class*="css"], .stMarkdown, p, div {
    font-family: 'HCR Dotum', '함초롱돋움', 'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif;
}

/* 기본 상단 헤더 숨김 */
header[data-testid="stHeader"] {
    display: none !important;
}

/* ⭐ 화면에 있는 모든 창의 바깥쪽 좌우 여백을 기존 대비 50% 축소 (0.65rem -> 0.32rem) */
.main .block-container, 
[data-testid="stAppViewBlockContainer"], 
.block-container {
    max-width: 100% !important;
    width: 100% !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-top: 0.6rem !important;
    padding-bottom: 1.0rem !important;
    padding-left: 0.32rem !important;
    padding-right: 0.32rem !important;
    box-sizing: border-box !important;
}

div[data-testid="stVerticalBlock"] {
    gap: 0.22rem !important;
}
div[data-testid="stElementContainer"],
.element-container {
    margin-top: 0px !important;
    margin-bottom: 1px !important;
}

/* 저장소 배경 이미지 적용 및 투명도 제로(100% 완전 불투명) */
.stApp {
    background-color: #162b52;
    background-size: cover !important;
    background-position: center top !important;
    background-attachment: fixed !important;
    opacity: 1 !important;
}

/* ⭐ 제일 위 페이지 타이틀 창: 내부 상하 여백 80% 축소 */
.title-container {
    width: 100% !important;
    box-sizing: border-box !important;
    margin-top: 4px !important;
    margin-bottom: 8px !important;
    padding: 1.5px 8px 1.5px 2px !important;
    border: 2px solid #b45309 !important;
    border-bottom: 3px solid #8c2d19 !important;
    background: #facc15 !important;
    background-color: #facc15 !important;
    opacity: 1 !important;
    border-radius: 6px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.25);
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    justify-content: flex-start !important;
    gap: 6px !important;
}

.title-left-icon {
    flex-shrink: 0 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    margin: 0px !important;
    padding: 0px !important;
}

.title-right-content {
    flex: 1 !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
    margin: 0px !important;
    padding: 0px !important;
}

.title-clan-name {
    font-size: 13pt;
    font-weight: 800;
    color: #1e3c72;
    line-height: 1.12;
    letter-spacing: -0.3px;
    margin-top: 0px;
    margin-bottom: 1px;
}

.title-main-center {
    font-size: 16pt;
    font-weight: 800;
    color: #111827;
    text-align: center;
    width: 100%;
    margin-top: 0px;
    margin-bottom: 1px;
    line-height: 1.12;
    letter-spacing: -0.5px;
}

.title-sub-bottom {
    font-size: 10.2pt;
    font-weight: 700;
    color: #1e293b;
    text-align: center;
    width: 100%;
    margin-top: 0px;
    margin-bottom: 0px;
    line-height: 1.18;
    letter-spacing: -0.2px;
}

/* 안내 문구 라인 */
.compact-search-header {
    width: 100% !important;
    text-align: center !important;
    margin-top: 6px !important;
    margin-bottom: 2px !important;
    padding: 1px 2px !important;
    line-height: 1.3 !important;
}
.compact-search-subtitle {
    font-size: 10.8pt !important;
    font-weight: 700 !important;
    color: #fef08a !important;
    text-align: center !important;
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    line-height: 1.3 !important;
    letter-spacing: -0.3px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* ⭐ 검색 유형 라디오 라인 (윗줄: 인명·계보·단어 검색 / 아랫줄: 분파도·문화재·추모 글 조회 동일 형식) */
div[data-testid="stRadio"] {
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}
div[role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    justify-content: center !important;
    gap: 12px !important;
    width: 100% !important;
    margin-top: 0px !important;
    margin-bottom: 0px !important;
    padding: 1px 2px !important;
    overflow-x: auto !important;
}
div[role="radiogroup"] > label,
div[role="radiogroup"] label[data-baseweb="radio"] {
    white-space: nowrap !important;
    flex-shrink: 0 !important;
    margin-right: 2px !important;
    margin-bottom: 0px !important;
    padding: 1px 2px !important;
    line-height: 1.25 !important;
}
div[role="radiogroup"] label p,
div[role="radiogroup"] label span {
    font-size: 12.0pt !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    white-space: nowrap !important;
    line-height: 1.25 !important;
    letter-spacing: -0.3px !important;
    text-shadow: 0 1px 3px rgba(0,0,0,0.65);
}

/* ⭐ 인명·계보·단어 검색 라인과 분파도·문화재·추모 글 조회 라인 사이의 새 가로줄 */
.radio-row-divider {
    border: none !important;
    border-top: 1.5px solid rgba(255, 255, 255, 0.55) !important;
    margin: 3px 4px 3px 4px !important;
    padding: 0px !important;
    height: 0px !important;
}

/* 검색어 입력창: 테두리 흰색(3.5px solid #ffffff), 배경 청색, 세로 중앙 정렬 */
div[data-testid="stTextInput"] {
    margin-top: 2px !important;
    margin-bottom: 4px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}
div[data-testid="stTextInputRootElement"] {
    border: 3.5px solid #ffffff !important;
    border-radius: 8px !important;
    background-color: #1d4ed8 !important;
    box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
    box-sizing: border-box !important;
    overflow: hidden !important;
}
div[data-testid="stTextInput"] div[data-baseweb="input"],
div[data-testid="stTextInput"] div[data-baseweb="base-input"] {
    background-color: #1d4ed8 !important;
    border: none !important;
    box-shadow: none !important;
    height: 46px !important;
    min-height: 46px !important;
    display: flex !important;
    align-items: center !important;
}
div[data-testid="stTextInput"] input {
    background-color: #1d4ed8 !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    caret-color: #ffffff !important;
    font-weight: 800 !important;
    font-size: 13.5pt !important;
    padding: 0px 12px !important;
    height: 100% !important;
    line-height: 46px !important;
    vertical-align: middle !important;
}
div[data-testid="stTextInput"] input::placeholder {
    color: #cbd5e1 !important;
    -webkit-text-fill-color: #cbd5e1 !important;
    font-size: 9.6pt !important;
    font-weight: 400 !important;
    opacity: 1 !important;
    line-height: 46px !important;
    vertical-align: middle !important;
}
div[data-testid="stTextInput"] input::-webkit-input-placeholder {
    color: #cbd5e1 !important;
    -webkit-text-fill-color: #cbd5e1 !important;
    font-size: 9.6pt !important;
    font-weight: 400 !important;
    opacity: 1 !important;
    line-height: 46px !important;
    vertical-align: middle !important;
}

/* 검색 결과 없을 때 안내 문구: 글자 색상 흰색(#ffffff) */
.no-result-notice {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: rgba(30, 58, 138, 0.88) !important;
    border: 1.5px solid #ffffff !important;
    color: #ffffff !important;
    padding: 9px 12px !important;
    border-radius: 6px !important;
    font-size: 10.8pt !important;
    font-weight: 500 !important;
    margin-top: 4px !important;
    margin-bottom: 6px !important;
    line-height: 1.45 !important;
    text-align: center !important;
    text-shadow: 0 1px 2px rgba(0,0,0,0.6) !important;
}

/* ⭐ '초기 검색 인물로 복귀' 버튼: 바탕색 녹색(#15803d), 글씨체 흰색(#ffffff) */
div[data-testid="stButton"] > button {
    background-color: #15803d !important;
    background: #15803d !important;
    color: #ffffff !important;
    border: 2px solid #ffffff !important;
    border-radius: 6px !important;
    font-weight: 800 !important;
    font-size: 11pt !important;
    padding: 6px 12px !important;
    box-shadow: 0 2px 5px rgba(0,0,0,0.3) !important;
}
div[data-testid="stButton"] > button:hover,
div[data-testid="stButton"] > button:active,
div[data-testid="stButton"] > button:focus {
    background-color: #166534 !important;
    background: #166534 !important;
    color: #ffffff !important;
    border-color: #fef08a !important;
}
div[data-testid="stButton"] > button *,
div[data-testid="stButton"] > button p,
div[data-testid="stButton"] > button span {
    color: #ffffff !important;
    font-weight: 800 !important;
    font-size: 11pt !important;
}

/* 복수 후보 안내 배너 */
.multi-ancestor-banner {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #eff6ff !important;
    border: 1.5px solid #3b82f6 !important;
    color: #1d4ed8 !important;
    padding: 7px 10px;
    border-radius: 6px;
    font-size: 11pt !important;
    font-weight: 400 !important;
    margin-top: 3px !important;
    margin-bottom: 5px !important;
    line-height: 1.45 !important;
}

/* 드롭다운 명부 */
div[data-testid="stSelectbox"] {
    margin-top: 2px !important;
    margin-bottom: 5px !important;
}
div[data-testid="stSelectbox"] *,
div[data-testid="stSelectbox"] div[data-baseweb="select"],
div[data-testid="stSelectbox"] div[data-baseweb="select"] div,
div[data-testid="stSelectbox"] div[data-baseweb="select"] span {
    font-size: 11.5pt !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    line-height: 1.4 !important;
}
ul[data-testid="stSelectboxVirtualDropdown"] li,
ul[data-testid="stSelectboxVirtualDropdown"] li *,
div[role="listbox"] li,
div[role="listbox"] li *,
div[role="option"],
div[role="option"] * {
    font-size: 11.5pt !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    line-height: 1.45 !important;
}

/* 검색 결과 인물 카드 */
.result-card {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #ffffff;
    border: 2px solid #1e3c72;
    border-top: 4px solid #1e3c72;
    border-radius: 8px;
    padding: 10px 12px;
    margin-top: 4px !important;
    margin-bottom: 8px !important;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}
.result-card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 5px;
    margin-bottom: 7px;
    line-height: 1.35;
}
.result-name-big {
    font-size: 15.5pt;
    font-weight: 800;
    color: #000000;
    display: flex;
    align-items: center;
    gap: 8px;
}
.result-badge {
    font-size: 11.5pt;
    font-weight: 800;
    background: #eff6ff;
    color: #1d4ed8;
    padding: 2px 8px;
    border-radius: 8px;
    border: 1px solid #bfdbfe;
}
/* ⭐ 모든 족보 페이지 표시: '시직공파보' 줄바꿈 '1권 433p 5단' 형식 + 글자크기 1p 축소 */
.result-book-loc {
    font-size: 9.5pt !important;
    font-weight: 800 !important;
    color: #1e293b !important;
    background: #f1f5f9 !important;
    padding: 2px 8px !important;
    border-radius: 6px !important;
    border: 1px solid #cbd5e1 !important;
    text-align: right !important;
    line-height: 1.22 !important;
    flex-shrink: 0 !important;
}

/* 한글 약력 컨테이너 */
.meta-para-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 8px 11px;
    margin-bottom: 8px;
    line-height: 1.52 !important;
    font-size: 10.5pt;
    color: #0f172a;
}
.meta-para-item {
    margin-bottom: 4px;
    line-height: 1.52 !important;
    word-break: keep-all;
    color: #0f172a;
}
.meta-para-label {
    font-weight: 700;
    color: #1e3c72;
    margin-right: 4px;
}
.meta-para-value {
    color: #1e293b;
}

/* 한문 약력 원문 전용 컨테이너 */
.hanmun-raw-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #fcfcfd;
    border: 1px solid #cbd5e1;
    border-left: 4px solid #1e3c72;
    border-radius: 6px;
    padding: 8px 12px;
    margin-top: 2px !important;
    margin-bottom: 2px !important;
    line-height: 1.55 !important;
    font-size: 10.5pt;
    color: #1e293b;
}

.red-correction {
    color: #dc2626 !important;
    font-weight: 700;
    background-color: #fef2f2;
    padding: 0px 3px;
    border-radius: 3px;
    border-bottom: 1px dashed #f87171;
}

.search-word-red {
    color: #ff0000 !important;
    font-weight: 700 !important;
    background-color: #fee2e2;
    padding: 0px 4px;
    border-radius: 3px;
    border: 1px solid #f87171;
}

/* ============================================================================
   ⭐ 계보도 검색 결과창 스타일
   ============================================================================ */
.tree-diagram-wrapper {
    width: 100% !important;
    box-sizing: border-box !important;
    background: #ffffff !important;
    border: 2px solid #1e3c72 !important;
    border-radius: 8px !important;
    padding: 8px 8px !important;
    margin-top: 2px !important;
    margin-bottom: 6px !important;
}
.tree-header-bar {
    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    border-bottom: 2px solid #e2e8f0 !important;
    padding-bottom: 5px !important;
    margin-bottom: 6px !important;
    gap: 8px !important;
}
.tree-level-row {
    display: flex !important;
    align-items: stretch !important;
    margin-bottom: 0px !important;
    padding: 4px 5px !important;
    background: #ffffff !important;
    background-color: #ffffff !important;
    border-radius: 5px !important;
    border: 1px solid #cbd5e1 !important;
    line-height: 1.3 !important;
}
.tree-level-row-self {
    display: flex !important;
    align-items: stretch !important;
    margin-bottom: 0px !important;
    padding: 4px 5px !important;
    background: #ffffff !important;
    background-color: #ffffff !important;
    border-radius: 5px !important;
    border: 1.5px solid #1d4ed8 !important;
    line-height: 1.3 !important;
}
.tree-connector-line {
    text-align: left !important;
    padding-left: 112px !important;
    color: #1e3c72 !important;
    font-weight: 900 !important;
    font-size: 8.5pt !important;
    line-height: 0.85 !important;
    margin: -1px 0px !important;
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}
.tree-gen-badge {
    min-width: 58px;
    width: 58px;
    font-weight: 800 !important;
    color: #1e3c72 !important;
    font-size: 8.6pt;
    padding: 2px 3px;
    background: #ffffff !important;
    background-color: #ffffff !important;
    border: 1px solid #94a3b8 !important;
    border-radius: 4px;
    text-align: center;
    margin-right: 5px;
    line-height: 1.18;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
}
.tree-gen-badge-self {
    min-width: 58px;
    width: 58px;
    font-weight: 800 !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    font-size: 8.6pt;
    padding: 2px 3px;
    background: #1d4ed8 !important;
    background-color: #1d4ed8 !important;
    border: 1px solid #1e3a8a !important;
    border-radius: 4px;
    text-align: center;
    margin-right: 5px;
    line-height: 1.18;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
}

.tree-split-wrapper {
    display: flex !important;
    flex-direction: row !important;
    align-items: flex-start !important;
    justify-content: flex-start !important;
    gap: 5px !important;
    flex: 1 !important;
    width: 100% !important;
    background: #ffffff !important;
}
.tree-left-male-zone {
    flex: 1 1 50% !important;
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: wrap !important;
    justify-content: flex-start !important;
    gap: 3px !important;
    align-items: center !important;
    align-content: flex-start !important;
    min-height: 22px !important;
}
.tree-right-female-zone {
    flex: 1 1 50% !important;
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: wrap !important;
    justify-content: flex-start !important;
    gap: 3px !important;
    align-items: center !important;
    align-content: flex-start !important;
    border-left: 1px dashed #cbd5e1 !important;
    padding-left: 5px !important;
    min-height: 22px !important;
}

.tree-cousin-steps-container {
    display: flex !important;
    flex-direction: column !important;
    flex: 1 !important;
    width: 100% !important;
    gap: 4px !important;
    background: #ffffff !important;
}
.tree-cousin-step {
    display: flex !important;
    flex-direction: row !important;
    align-items: stretch !important;
    justify-content: flex-start !important;
    width: 100% !important;
    gap: 4px !important;
    background: #ffffff !important;
}
.tree-father-vert-box {
    width: 22px !important;
    min-width: 22px !important;
    max-width: 22px !important;
    background-color: #f8fafc !important;
    border: 1px solid #64748b !important;
    border-radius: 4px !important;
    color: #0f172a !important;
    font-size: 8.2pt !important;
    font-weight: 800 !important;
    line-height: 1.08 !important;
    padding: 3px 1px !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
    flex-shrink: 0 !important;
    box-sizing: border-box !important;
}
.tree-children-group-box {
    flex: 1 !important;
    width: 100% !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 4px !important;
    padding: 3px 4px !important;
    background-color: #ffffff !important;
    display: flex !important;
    align-items: center !important;
    box-sizing: border-box !important;
}

.tree-node-box {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    text-align: left !important;
    text-decoration: none !important;
    cursor: pointer !important;
    width: 108px !important;
    min-width: 108px !important;
    max-width: 108px !important;
    box-sizing: border-box !important;
    border-radius: 4px !important;
    padding: 2px 4px !important;
    font-size: 8.3pt !important;
    letter-spacing: -0.35px !important;
    line-height: 1.22 !important;
    font-weight: 700 !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    box-shadow: none !important;
}
.tree-node-box * {
    border: none !important;
    background: transparent !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
}

a.tree-node-direct {
    background-color: #facc15 !important;
    background: #facc15 !important;
    border: 1px solid #b45309 !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 800 !important;
}
a.tree-node-collateral {
    background-color: #fef08a !important;
    background: #fef08a !important;
    border: 1px solid #eab308 !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 700 !important;
}
a.tree-node-chulgye {
    background-color: #fefce8 !important;
    background: #fefce8 !important;
    border: 1px solid #fde047 !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 700 !important;
}
a.tree-node-self {
    background-color: #1d4ed8 !important;
    background: #1d4ed8 !important;
    border: 1px solid #1e3a8a !important;
    color: #ffffff !important;
    -webkit-text-fill-color: #ffffff !important;
    font-weight: 800 !important;
}
a.tree-node-male {
    background-color: #dbeafe !important;
    background: #dbeafe !important;
    border: 1px solid #60a5fa !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 700 !important;
}
a.tree-node-female {
    background-color: #ffffff !important;
    background: #ffffff !important;
    border: 1px solid #94a3b8 !important;
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 700 !important;
}

.heritage-sub-container {
    width: 100% !important;
    box-sizing: border-box !important;
    background-color: #15803d !important;
    border: 2px solid #14532d !important;
    border-radius: 8px;
    padding: 9px 12px;
    margin-top: 6px;
    margin-bottom: 6px;
    box-shadow: 0 3px 8px rgba(0,0,0,0.22);
}
.heritage-sub-header {
    font-size: 11.2pt;
    font-weight: 800;
    color: #ffffff !important;
    margin-bottom: 6px;
    line-height: 1.4;
    text-shadow: 0 1px 2px rgba(0,0,0,0.5);
}
.heritage-sub-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 6px;
}
.heritage-sub-card {
    background: #ffffff;
    border: 1px solid #dcfce7;
    border-radius: 6px;
    padding: 6px 8px;
    text-decoration: none !important;
    display: block;
}
.heritage-sub-card:hover {
    background: #f0fdf4;
    border-color: #fef08a;
}
.h-card-name {
    font-size: 10pt;
    font-weight: 700;
    color: #14532d;
    margin-bottom: 2px;
    line-height: 1.35;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.h-card-desc {
    font-size: 8.8pt;
    color: #334155;
    line-height: 1.35;
}

@media (max-width: 768px) {
    html,
    body,
    #root {
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        box-sizing: border-box !important;
    }

    .stApp,
    [data-testid="stAppViewContainer"],
    section.main,
    [data-testid="stAppViewBlockContainer"],
    [data-testid="stMain"] {
        width: 100% !important;
        max-width: 100% !important;
        margin-left: 0 !important;
        margin-right: 0 !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        box-sizing: border-box !important;
    }

    [data-testid="stMainBlockContainer"],
    .main .block-container,
    .block-container {
        position: relative !important;
        left: calc((100vw - 100%) / 2) !important;
        width: 100% !important;
        max-width: 100% !important;
        margin-left: auto !important;
        margin-right: auto !important;
        padding-left: 6px !important;
        padding-right: 6px !important;
        box-sizing: border-box !important;
    }

    [data-testid="stMarkdownContainer"],
    .stMarkdown {
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }

    .tree-gen-badge,
    .tree-gen-badge-self {
        min-width: 50px !important;
        width: 50px !important;
        font-size: 8.1pt !important;
        margin-right: 4px !important;
        padding: 2px 2px !important;
    }

    .tree-connector-line {
        padding-left: 98px !important;
    }

    .tree-father-vert-box {
        width: 20px !important;
        min-width: 20px !important;
        max-width: 20px !important;
        font-size: 7.9pt !important;
    }

    .tree-node-box {
        width: 98px !important;
        min-width: 98px !important;
        max-width: 98px !important;
        font-size: 7.9pt !important;
        letter-spacing: -0.45px !important;
        padding: 2px 3px !important;
    }
}
</style>
"""

render_html(CUSTOM_CSS)

# =============================================================================
# 2. 데이터 로드 및 초고속(0.05초) 메모리 인덱싱
# =============================================================================
VOL_ORDER_MAP = {"A": 1, "B": 2, "C": 3, "D": 4}

KO_SEARCH_PATTERNS = [re.compile(p) for p in [
    r'자는\s*([^\.,\n\s]+)',
    r'초자는?\s*([^\.,\n\s]+)',
    r'호는\s*([^\.,\n\s]+)',
    r'일휘는?\s*([^\.,\n\s]+)',
    r'초휘는?\s*([^\.,\n\s]+)',
    r'초명은?\s*([^\.,\n\s]+)',
    r'통명은?\s*([^\.,\n\s]+)',
]]

CN_SEARCH_PATTERNS = [re.compile(p) for p in [
    r'字\s*([^\.,\n\s]+)',
    r'初字\s*([^\.,\n\s]+)',
    r'號\s*([^\.,\n\s]+)',
    r'一諱\s*([^\.,\n\s]+)',
    r'初諱\s*([^\.,\n\s]+)',
    r'初名\s*([^\.,\n\s]+)',
    r'通名\s*([^\.,\n\s]+)',
]]

def parse_gid_tuple(gid):
    m = re.match(r"^([A-Z])-(\d+)-(\d+)-(\d+)-(\d+)$", str(gid).strip())
    if m:
        return m.group(1), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5))
    return "", 0, 0, 0, 0

def format_book_location_html(gid):
    vol, page, tier, _, _ = parse_gid_tuple(gid)
    if not vol:
        return str(gid).strip()
    vol_num = "1권" if vol in ("A", "B", "C") else ("2권" if vol == "D" else f"{vol}권")
    return f"시직공파보<br>{vol_num} {page}p {tier}단"

def get_sibling_sort_key(row, all_sibling_rels=None):
    rel = str(row.get("관계", "")).strip()
    vol, page, tier, seq, _ = parse_gid_tuple(row.get("고유ID", ""))
    vol_idx = VOL_ORDER_MAP.get(vol, 9)

    if "녀" in rel:
        gender_rank = 1
        m = re.match(r"^(\d+)녀$", rel)
        rel_num = float(m.group(1)) if m else 99.0
    else:
        gender_rank = 0
        m = re.match(r"^(\d+)자$", rel)
        if m:
            rel_num = float(m.group(1))
        elif rel in ("계자", "승사손"):
            rel_num = 1.5 if (all_sibling_rels and "1자" in all_sibling_rels) else 1.0
        else:
            rel_num = 99.0

    return (gender_rank, rel_num, vol_idx, page, tier, seq)

def sort_siblings_df(sib_df):
    if sib_df is None or len(sib_df) <= 1:
        return sib_df
    rels = set(sib_df["관계"].astype(str).str.strip().tolist())
    recs = sib_df.to_dict("records")
    recs.sort(key=lambda r: get_sibling_sort_key(r, rels))
    return pd.DataFrame(recs)

@st.cache_resource(show_spinner=False)
def load_and_preprocess_genealogy():
    raw_df = None
    status_msg = ""

    for cand_db in DB_PATHS:
        if os.path.exists(cand_db):
            try:
                conn = sqlite3.connect(f"file:{cand_db}?immutable=1", uri=True, check_same_thread=False)
                raw_df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
                conn.close()
                if not raw_df.empty:
                    status_msg = "성공 (SQLite DB)"
                    break
            except Exception:
                try:
                    conn = sqlite3.connect(cand_db, check_same_thread=False)
                    raw_df = pd.read_sql("SELECT * FROM genealogy", conn).fillna("")
                    conn.close()
                    if not raw_df.empty:
                        status_msg = "성공 (SQLite DB)"
                        break
                except Exception:
                    raw_df = None

    if (raw_df is None or raw_df.empty) and os.path.exists(DATA_FILE):
        try:
            wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
            sheet_name = "족보전사결과" if "족보전사결과" in wb.sheetnames else wb.sheetnames[0]
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if not rows:
                return None, None, None, "시트에 데이터가 없습니다."
            header = [str(c).strip() if c is not None else f"Col{idx}" for idx, c in enumerate(rows[0])]
            raw_df = pd.DataFrame(rows[1:], columns=header).fillna("")
            status_msg = "성공 (Excel)"
        except Exception as e:
            return None, None, None, str(e)

    if raw_df is None or raw_df.empty:
        return None, None, None, f"데이터 파일({DB_FILE} 또는 {DATA_FILE})을 찾을 수 없습니다."

    records = raw_df.to_dict("records")
    by_name = {}
    search_blobs = []

    for r in records:
        vol, page, tier, seq, gen = parse_gid_tuple(r.get("고유ID", ""))
        r["_vol"] = vol
        r["_page"] = page
        r["_tier"] = tier
        r["_seq"] = seq
        r["_gen"] = gen
        nm = str(r.get("인명", "")).strip()
        by_name.setdefault(nm, []).append(r)

        ko_name = re.sub(r'[^가-힣]+', '', nm)
        cn_name = re.sub(r'[^一-龥]+', '', nm)
        tokens = [nm, ko_name, cn_name]
        if ko_name and cn_name:
            tokens.append(f"{ko_name}_{cn_name}")
            tokens.append(f"{ko_name}{cn_name}")
        if ko_name:
            tokens.append('안' + ko_name)
        if cn_name:
            tokens.append('安' + cn_name)

        fl_ko = str(r.get('한글 약력 정리', '')).split('\n')[0]
        for pat in KO_SEARCH_PATTERNS:
            for m in pat.finditer(fl_ko):
                tokens.append(m.group(1).strip())

        fl_cn = str(r.get('한문 약력 정리', '')).split('\n')[0]
        for pat in CN_SEARCH_PATTERNS:
            for m in pat.finditer(fl_cn):
                tokens.append(m.group(1).strip())

        search_blobs.append("\n".join(t for t in tokens if t))

    father_gids = []
    for r in records:
        f_name = str(r.get("부친", "")).strip()
        if not f_name or f_name == "覯구" or f_name == "-":
            r["_father_gid"] = ""
            father_gids.append("")
            continue
        cands_all = by_name.get(f_name, [])
        if not cands_all:
            r["_father_gid"] = ""
            father_gids.append("")
            continue
        if len(cands_all) == 1:
            fg = str(cands_all[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        c_gen = r["_gen"]
        cands = [c for c in cands_all if c["_gen"] == c_gen - 1] or cands_all
        if len(cands) == 1:
            fg = str(cands[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        gf_name = str(r.get("조부", "")).strip()
        if gf_name and gf_name != "-":
            cands_gf = [c for c in cands if str(c.get("부친", "")).strip() == gf_name]
            if len(cands_gf) == 1:
                fg = str(cands_gf[0]["고유ID"])
                r["_father_gid"] = fg
                father_gids.append(fg)
                continue
            elif len(cands_gf) > 1:
                cands = cands_gf

        c_vol, c_page = r["_vol"], r["_page"]
        same_vol = [c for c in cands if c["_vol"] == c_vol and c["_page"] <= c_page]
        if same_vol:
            same_vol.sort(key=lambda x: (-x["_page"], x["_seq"]))
            fg = str(same_vol[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        non_empty = [c for c in cands if ("생" in str(c.get("한글 약력 정리", "")) or "배 :" in str(c.get("한글 약력 정리", "")))]
        if len(non_empty) == 1:
            fg = str(non_empty[0]["고유ID"])
            r["_father_gid"] = fg
            father_gids.append(fg)
            continue

        fg = str(cands[0]["고유ID"])
        r["_father_gid"] = fg
        father_gids.append(fg)

    raw_df["_father_gid"] = father_gids
    raw_df["_search_blob"] = search_blobs

    gid_to_row = {str(r["고유ID"]).strip(): r for r in records}
    children_by_fgid = {}
    for r in records:
        fg = r["_father_gid"]
        if fg:
            children_by_fgid.setdefault(fg, []).append(r)
        elif str(r.get("부친", "")).strip() == "覯구":
            children_by_fgid.setdefault("__ROOT_GU__", []).append(r)

    for fg, ch_list in children_by_fgid.items():
        if len(ch_list) > 1:
            rels = {str(x.get("관계", "")).strip() for x in ch_list}
            ch_list.sort(key=lambda x: get_sibling_sort_key(x, rels))

    return raw_df, gid_to_row, children_by_fgid, status_msg

df, gid_to_row, children_by_fgid, load_msg = load_and_preprocess_genealogy()

# =============================================================================
# 3. 보조 유틸리티 함수
# =============================================================================
EXAM_GROUP_KEYWORDS = [
    "과거", "문과", "급제", "사마", "별시", "대과", "무과", "잡과", "장원",
    "진사", "생원", "향시", "회시", "복시", "현랑과", "현량과", "증광시",
    "전시", "식년시", "알성시", "정시", "춘당대시", "갑과", "을과", "병과"
]

LITERARY_GROUP_KEYWORDS = [
    "문집", "유집", "유고", "실기", "서집록", "저서", "논문"
]

def sort_search_candidates_df(cand_df):
    if cand_df is None or len(cand_df) <= 1:
        return cand_df
    recs = cand_df.to_dict("records")
    def cand_key(r):
        rel = str(r.get("관계", "")).strip()
        gender_rank = 1 if "녀" in rel else 0
        vol, page, tier, seq, _ = parse_gid_tuple(r.get("고유ID", ""))
        vol_idx = VOL_ORDER_MAP.get(vol, 9)
        return (gender_rank, vol_idx, page, tier, seq)
    recs.sort(key=cand_key)
    return pd.DataFrame(recs)

def fast_search_by_name(all_df, q):
    q = q.strip()
    if not q:
        return all_df.iloc[0:0]
    q_alt = q[1:] if q.startswith('안') and len(q) > 1 else q
    pat = re.escape(q) if q == q_alt else f"{re.escape(q)}|{re.escape(q_alt)}"
    mask = all_df["_search_blob"].str.contains(pat, na=False, regex=True)
    return sort_search_candidates_df(all_df[mask].reset_index(drop=True))

def extract_generation_info(gid):
    m = re.search(r"-(\d+)$", str(gid).strip())
    if m:
        gen = int(m.group(1))
        wangwi_desc_gen = gen - 24
        desc_str = "완귀공 본인" if wangwi_desc_gen == 0 else f"완귀공의 {wangwi_desc_gen}대손"
        return gen, f"{gen}세  {desc_str}"
    return None, "세수 미상"

def split_name(name_raw):
    s = str(name_raw).strip()
    m = re.match(r"^([一-龥]+)?([가-힣]+)?$", s)
    if m and m.group(1) and m.group(2):
        return m.group(2), m.group(1)
    if re.fullmatch(r"^[가-힣]+$", s):
        return s, ""
    if re.fullmatch(r"^[一-龥]+$", s):
        return "", s
    return s, ""

def clean_ko_name(raw_val):
    s = str(raw_val).strip()
    ko, _ = split_name(s)
    if ko:
        return ko
    cleaned = re.sub(r"[一-龥]+", "", s).strip()
    return cleaned if cleaned else s

def format_ko_hanja_attached(raw_val):
    s = str(raw_val).strip()
    ko, hj = split_name(s)
    if ko and hj:
        return f"{ko}_{hj}"
    return s

def get_adoption_info(row):
    rel = str(row.get("관계", "")).strip()
    ko_bio = str(row.get("한글 약력 정리", "")).strip()
    ko_raw = str(row.get("한글 약력 원문", "")).strip()
    cn_raw = str(row.get("한문 약력 원문", "")).strip()
    father_ko = clean_ko_name(row.get("부친", ""))

    if rel == "계자" or "생부 " in ko_bio:
        m = re.search(r"생부\s*([^\s의\.]+)의?", ko_bio)
        saengbu_ko = clean_ko_name(m.group(1)) if m else ""
        return "입계", father_ko, saengbu_ko

    if "양부" in ko_bio or "출계" in ko_raw or "出系" in cn_raw:
        m = re.search(r"양부\s*(.+?)의\s*계자", ko_bio)
        yangbu_ko = clean_ko_name(m.group(1)) if (m and m.group(1).strip()) else ""
        if not yangbu_ko:
            m2 = re.search(r"([^\s]+)의\s*후사로\s*출계|출계\s*([^\s]+)", ko_raw)
            if m2:
                yangbu_ko = clean_ko_name(m2.group(1) or m2.group(2))
        return "출계", father_ko, yangbu_ko

    return "일반", father_ko, ""

def format_tree_node_label(row):
    base_name = format_ko_hanja_attached(row.get("인명", ""))
    adopt_type, _, _ = get_adoption_info(row)
    if adopt_type == "입계":
        return f"{base_name}&nbsp;&nbsp;&nbsp;계자"
    elif adopt_type == "출계":
        return f"{base_name}&nbsp;&nbsp;&nbsp;출계"
    return base_name

def get_tree_node_class(row, is_self=False, is_direct=False, is_ancestor_row=False):
    if is_self:
        return "tree-node-box tree-node-self"
    if is_direct:
        return "tree-node-box tree-node-direct"
    rel = str(row.get("관계", "")).strip()
    if "녀" in rel:
        return "tree-node-box tree-node-female"
    adopt_type, _, _ = get_adoption_info(row)
    if adopt_type == "출계":
        return "tree-node-box tree-node-chulgye"
    if is_ancestor_row:
        return "tree-node-box tree-node-collateral"
    return "tree-node-box tree-node-male"

def format_candidate_label(row):
    person_combined = format_ko_hanja_attached(row.get("인명", ""))
    gfather_ko = clean_ko_name(row.get("조부", ""))
    rel = str(row.get("관계", "")).strip()
    adopt_type, father_ko, other_father_ko = get_adoption_info(row)

    if adopt_type == "입계":
        rel_str = rel if rel == "계자" else f"{rel}(계자)"
        parts = [f"{person_combined}: {rel_str}"]
        if father_ko:
            parts.append(f"부 {father_ko}")
        if other_father_ko:
            parts.append(f"생부 {other_father_ko}")
        if gfather_ko:
            parts.append(f"조부 {gfather_ko}")
        return ", ".join(parts)

    elif adopt_type == "출계":
        rel_str = f"{rel}(출계)" if "출계" not in rel else rel
        parts = [f"{person_combined}: {rel_str}"]
        if father_ko:
            parts.append(f"부 {father_ko}")
        if other_father_ko:
            parts.append(f"양부 {other_father_ko}")
        if gfather_ko:
            parts.append(f"조부 {gfather_ko}")
        return ", ".join(parts)

    else:
        parts = [f"{person_combined}: {rel}"]
        if father_ko:
            parts.append(f"부 {father_ko}")
        if gfather_ko:
            parts.append(f"조부 {gfather_ko}")
        return ", ".join(parts)

def highlight_corrections_html(text, search_term=None):
    if not text:
        return ""
    keywords = [
        "1699년", "1699", "一六九九", "1626년", "1626", "一六二六",
        "1558년", "1558", "一五五八", "1816년", "1816", "一八一六年",
        "1814년", "1814", "一八一四", "1771년", "1771", "一七七一",
        "1740년", "1740", "一七四〇", "1750년", "1750", "一七五〇",
        "1775년", "1775", "一七七5", "1898년", "1898", "一八九八",
        "1910년", "1910", "一九一〇", "1863년", "1863", "一八六三",
        "1850년", "1850", "一八五〇", "1870년", "1870", "一八七〇",
        "순종", "純宗", "헌종", "憲宗", "철종", "哲宗", "고종", "高宗",
        "기해", "己亥", "무신", "戊申", "자좌 오향", "자좌 오", "子坐 午向", "子坐 午",
        "내등말", "內嶝末", "오보허", "五步許", "북록 묘좌", "北麓 卯坐",
        "신재 주세붕", "慎齋 周世鵬", "병와 이형상", "瓶窩 李衡祥",
        "정헌 이종상", "定軒 李種祥", "순암 정복", "順庵 鼎福", "척암 김도화", "拓庵 金道和"
    ]
    formatted = str(text)
    for kw in keywords:
        if kw in formatted:
            formatted = formatted.replace(kw, f'<span class="red-correction">{kw}</span>')

    if search_term:
        if isinstance(search_term, (list, tuple, set)):
            terms = [str(t).strip() for t in search_term if str(t).strip()]
            terms.sort(key=len, reverse=True)
            if terms:
                pat = "|".join(re.escape(t) for t in terms)
                formatted = re.sub(
                    f"({pat})",
                    r'<span class="search-word-red">\1</span>',
                    formatted,
                    flags=re.IGNORECASE
                )
        elif str(search_term).strip():
            sterm = str(search_term).strip()
            formatted = re.sub(
                f"({re.escape(sterm)})",
                r'<span class="search-word-red">\1</span>',
                formatted,
                flags=re.IGNORECASE
            )
    return formatted

def extract_munjip_info(ko_bio):
    for line in str(ko_bio).split("\n"):
        if line.startswith("문집 :"):
            return line.replace("문집 :", "").strip()
        if line.startswith("학문·인품·업적 :"):
            c = line.replace("학문·인품·업적 :", "").strip()
            if "설서 완귀선생 유사" in c:
                return "설서 완귀선생 유사(說書 完龜先生 遺事)."
            if "문집이 있어 세상에 행해지며" in c:
                return "문집이 있어 세상에 행해짐."
            if "실기 상하 두 책이 있다" in c:
                return "실기(實記) 상·하 두 책이 있음."
            if "유고와 화부재강록이 있다" in c:
                return "유고(遺稿)와 화부재강록이 있음."
            if "고종 신축 1901년에 적변으로 유고가 소실되었다" in c:
                return "유고(고종 신축 1901년 적변으로 소실)."
            if "집안에 소장된 유고가 있" in c or "유고가 집안에 소장" in c:
                return "집안에 소장된 유고(遺稿)가 있음."
            if "유집이 있다" in c:
                return "유집(遺集)이 있음."
            if "학술논문집이 있다" in c:
                return "학술논문집이 있음."
            if "문집이 있다" in c:
                return "문집(文集)이 있음."
            if "유고가 있" in c:
                return "유고(遺稿)가 있음."
            if "서집록이 있다" in c:
                return "서집록(書集錄)이 있음."
            m = re.search(r"(저서[^\.]+있다)", c)
            if m:
                return m.group(1).strip() + "."
    return ""

SENTENCE_END_PATTERN = re.compile(
    r'(하였다|있었다|되었다|받았다|보임되었다|역임하였다|전하였다|없었다|추중하였다|존중하였다|칭송하였다|마쳤다|올랐다|돌아왔다|소실되었다|세웠다|받들었다|수훈하였다|길렀다|높았다|있도다|빛나리라|격절되었도다|만하도다|아름다움을 이었다|실려 있다|기록이다|것이다|있다|지냈다|찬술하였다|개찬하였다|배향되었다|행사한다|따님이다)\s+(?!(?:고|며|니|하고|하니|하며|하였으니|하였다|하더라도))'
)

def ensure_sentence_periods(text, with_linebreak=False):
    if not text:
        return ""
    s = str(text).strip()
    if not s or s == "-":
        return s
    s = SENTENCE_END_PATTERN.sub(r'\1. ', s)
    parts = [p.strip() for p in re.split(r'\.\s*', s) if p.strip()]
    formatted_parts = [p if p.endswith('.') else p + '.' for p in parts]
    if with_linebreak:
        return "<br>".join(formatted_parts)
    return " ".join(formatted_parts)

def parse_person_hangul_details(row):
    gid = row.get("고유ID", "")
    _, gen_str = extract_generation_info(gid)
    gfather_ko = clean_ko_name(row.get("조부", ""))
    adopt_type, father_ko, other_father_ko = get_adoption_info(row)

    if adopt_type == "입계" and other_father_ko:
        father_display = f"{father_ko}(생부 {other_father_ko})"
    elif adopt_type == "출계" and other_father_ko:
        father_display = f"{father_ko}(양부 {other_father_ko})"
    else:
        father_display = father_ko

    fg_parts = []
    if father_display and father_display != "-":
        fg_parts.append(f"부친 {father_display}")
    if gfather_ko and gfather_ko != "-":
        fg_parts.append(f"조부 {gfather_ko}")
    father_gfather_str = ", ".join(fg_parts)

    ko_bio = str(row.get("한글 약력 정리", ""))
    first_ko = ko_bio.split("\n")[0] if ko_bio else ""

    details = [
        ("세수", ensure_sentence_periods(gen_str)),
    ]
    if father_gfather_str:
        details.append(("부친·조부", ensure_sentence_periods(father_gfather_str)))

    m_ja = re.search(r'자는\s*([^\.,\n]+)', first_ko)
    m_ho = re.search(r'호는\s*([^\.,\n]+)', first_ko)
    jaho_parts = []
    if m_ja:
        jaho_parts.append(f"자 {m_ja.group(1).strip()}")
    if m_ho:
        jaho_parts.append(f"호 {m_ho.group(1).strip()}")
    if jaho_parts:
        details.append(("자·호", ensure_sentence_periods(", ".join(jaho_parts))))

    m = re.search(r'초자는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초자", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'초휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초휘", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'일휘는?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("일휘", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'통명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("통명", ensure_sentence_periods(m.group(1).strip())))
    m = re.search(r'초명은?\s*([^\.,\n]+)', first_ko)
    if m: details.append(("초명", ensure_sentence_periods(m.group(1).strip())))

    first_parts = [p.strip() for p in first_ko.split(".") if p.strip()]
    birth_parts = []
    death_parts = []
    other_first_parts = []
    for p in first_parts:
        if re.match(r"^(자는|호는|초자는?|초휘는?|일휘는?|통명은?|초명은?|생부\s|양부\s)", p):
            continue
        if p.endswith("생") or "년생" in p:
            birth_parts.append(p)
        elif "졸" in p or "향년" in p or "수 " in p or "별세" in p or "기일" in p:
            death_parts.append(p)
        elif p != "-":
            other_first_parts.append(p)

    if birth_parts:
        details.append(("생년", ensure_sentence_periods(". ".join(birth_parts))))
    if death_parts:
        details.append(("졸년", ensure_sentence_periods(". ".join(death_parts))))
    if other_first_parts:
        details.append(("신상", ensure_sentence_periods(". ".join(other_first_parts))))

    gwageo_val = ""
    hakmun_val = ""
    myo_list = []
    bae_list = []

    for line in ko_bio.split("\n"):
        if line.startswith("과거·관직 :"):
            gwageo_val = line.replace("과거·관직 :", "").strip()
        elif line.startswith("학문·인품·업적 :"):
            hakmun_val = line.replace("학문·인품·업적 :", "").strip()
        elif line.startswith("묘 :"):
            myo_list.append(line.replace("묘 :", "").strip())
        elif line.startswith("배 :") or line.startswith("배위 :"):
            bae_list.append(re.sub(r'^(배|배위)\s*:\s*', '', line).strip())

    if gwageo_val:
        details.append(("과거·관직", ensure_sentence_periods(gwageo_val)))

    munjip_val = extract_munjip_info(ko_bio)
    if munjip_val:
        details.append(("문집", ensure_sentence_periods(munjip_val)))

    if hakmun_val:
        details.append(("학력·인품·업적", ensure_sentence_periods(hakmun_val, with_linebreak=True)))

    for myo in myo_list:
        details.append(("묘", ensure_sentence_periods(myo)))
    for bae in bae_list:
        details.append(("배위", ensure_sentence_periods(bae)))

    return details

def format_hangul_only_paragraph_html(row, search_term=None):
    details = parse_person_hangul_details(row)
    lines_html = []

    for lbl, ko_val in details:
        val = str(ko_val).strip()
        if not val or val == "-":
            continue
        if search_term:
            val = highlight_corrections_html(val, search_term)
        lines_html.append(
            f"<div class='meta-para-item'><span class='meta-para-label'>• {lbl} :</span> <span class='meta-para-value'>{val}</span></div>"
        )

    return "".join(lines_html)

def render_person_result_card(target_row, search_term=None):
    combined_name = format_ko_hanja_attached(target_row['인명'])
    adopt_type, _, _ = get_adoption_info(target_row)
    rel_badge_text = f"{target_row['관계']}(출계)" if adopt_type == "출계" else str(target_row['관계'])
    book_loc_html = format_book_location_html(target_row.get('고유ID', ''))

    meta_paragraphs_html = format_hangul_only_paragraph_html(target_row, search_term=search_term)
    cn_raw = highlight_corrections_html(target_row.get('한문 약력 원문', ''), search_term=search_term)

    render_html(f"""
    <div id="search-result-anchor"></div>
    <div class="result-card">
        <div class="result-card-header">
            <div class="result-name-big">
                <span>{combined_name}</span>
                <span class="result-badge">{rel_badge_text}</span>
            </div>
            <div class="result-book-loc">{book_loc_html}</div>
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-bottom:4px; font-size:11pt;">
            📌 한글 약력
        </div>
        <div class="meta-para-container">
            {meta_paragraphs_html}
        </div>
        
        <div style="font-weight:800; color:#1e3c72; margin-top:6px; margin-bottom:4px; font-size:11pt;">
            📜 한문 약력 원문
        </div>
        <div class="hanmun-raw-container">
            {cn_raw if cn_raw else '수록된 한문 약력 원문이 없습니다.'}
        </div>
    </div>
    """)

def get_siblings_list_fast(person_row):
    f_gid = str(person_row.get("_father_gid", "")).strip()
    if f_gid:
        return children_by_fgid.get(f_gid, [person_row])
    f_name = str(person_row.get("부친", "")).strip()
    if f_name == "覯구":
        return children_by_fgid.get("__ROOT_GU__", [person_row])
    return [person_row]

def build_split_zones_html(male_htmls, female_htmls):
    left_content = "".join(male_htmls)
    right_content = "".join(female_htmls)
    return (
        f'<div class="tree-split-wrapper">'
        f'<div class="tree-left-male-zone">{left_content}</div>'
        f'<div class="tree-right-female-zone">{right_content}</div>'
        f'</div>'
    )

def render_tree_hierarchy(target_row, all_df, initial_gid=None):
    target_gid = str(target_row['고유ID']).strip()
    root_init_gid = str(initial_gid).strip() if initial_gid else target_gid
    target_name = target_row['인명']
    person_label = format_ko_hanja_attached(target_name)
    book_loc_html = format_book_location_html(target_gid)

    ancestors = []
    curr = target_row
    up = 1
    while True:
        f_gid = str(curr.get('_father_gid', '')).strip()
        if not f_gid:
            break
        f_row = gid_to_row.get(f_gid)
        if not f_row:
            break
        sibs = get_siblings_list_fast(f_row)
        ancestors.append({"level": up, "row": f_row, "siblings": sibs})
        curr = f_row
        up += 1
        if up > 30:
            break

    diagram_parts = [
        f'<div id="search-result-anchor"></div>'
        f'<div class="tree-diagram-wrapper">'
        f'<div class="tree-header-bar">'
        f'<div style="font-weight:800; color:#000000; font-size:12pt; line-height:1.35;">'
        f'🌳 계보도 [{person_label}]<br>'
        f'<span style="font-size:9.5pt; font-weight:700; color:#dc2626;">인물 박스 클릭시 새 계보 생성</span>'
        f'</div>'
        f'<div class="result-book-loc">{book_loc_html}</div>'
        f'</div>'
    ]
    
    for idx_a, anc in enumerate(reversed(ancestors)):
        lvl = anc['level']
        r = anc['row']
        g_num, _ = extract_generation_info(r['고유ID'])
        sibs = anc['siblings']
        
        if idx_a > 0:
            diagram_parts.append('<div class="tree-connector-line">▼</div>')

        if lvl == 1:
            lbl_badge = "부친"
        elif lvl == 2:
            lbl_badge = "조부"
        elif lvl == 3:
            lbl_badge = "증조"
        elif lvl == 4:
            lbl_badge = "고조"
        else:
            lbl_badge = f"{lvl}대조"

        if lvl >= 5:
            r_gid = str(r['고유ID']).strip()
            r_label = format_tree_node_label(r)
            n_cls = get_tree_node_class(r, is_self=False, is_direct=True, is_ancestor_row=True)
            node_tag = f'<a href="?tree_gid={r_gid}&init_gid={root_init_gid}" target="_self" class="{n_cls}">{r_label}</a>'
            split_html = build_split_zones_html([node_tag], [])
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}<br>{g_num}세</span>
                {split_html}
            </div>
            """)
        else:
            male_htmls = []
            female_htmls = []
            for s in sibs:
                is_direct = (s['고유ID'] == r['고유ID'])
                s_label = format_tree_node_label(s)
                s_rel = str(s['관계']).strip()
                s_gid = str(s['고유ID']).strip()
                n_cls = get_tree_node_class(s, is_self=False, is_direct=is_direct, is_ancestor_row=True)

                node_tag = f'<a href="?tree_gid={s_gid}&init_gid={root_init_gid}" target="_self" class="{n_cls}">{s_label}</a>'
                if "녀" in s_rel:
                    female_htmls.append(node_tag)
                else:
                    male_htmls.append(node_tag)

            split_html = build_split_zones_html(male_htmls, female_htmls)
            diagram_parts.append(f"""
            <div class="tree-level-row">
                <span class="tree-gen-badge">▲ {lbl_badge}<br>{g_num}세</span>
                {split_html}
            </div>
            """)

    if len(ancestors) > 0:
        diagram_parts.append('<div class="tree-connector-line">▼</div>')

    t_sibs = get_siblings_list_fast(target_row)
    tg_num, _ = extract_generation_info(target_row['고유ID'])

    t_male_htmls = []
    t_female_htmls = []
    for s in t_sibs:
        is_t = (s['고유ID'] == target_row['고유ID'])
        s_label = format_tree_node_label(s)
        s_rel = str(s['관계']).strip()
        s_gid = str(s['고유ID']).strip()
        n_cls = get_tree_node_class(s, is_self=is_t, is_direct=False, is_ancestor_row=False)

        node_tag = f'<a href="?tree_gid={s_gid}&init_gid={root_init_gid}" target="_self" class="{n_cls}">{s_label}</a>'
        if "녀" in s_rel:
            t_female_htmls.append(node_tag)
        else:
            t_male_htmls.append(node_tag)

    t_split_html = build_split_zones_html(t_male_htmls, t_female_htmls)
    diagram_parts.append(f"""
    <div class="tree-level-row-self">
        <span class="tree-gen-badge-self">● 본인<br>{tg_num}세</span>
        {t_split_html}
    </div>
    """)

    max_down = 4
    down_labels = {
        1: "자",
        2: "손",
        3: "증손",
        4: "현손"
    }
    current_parents = [target_row]
    for down_lvl in range(1, max_down + 1):
        cousin_step_htmls = []
        next_gen_rows = []

        for p_row in current_parents:
            p_gid = str(p_row['고유ID']).strip()
            sorted_children = children_by_fgid.get(p_gid, [])
            if not sorted_children:
                continue

            step_male_htmls = []
            step_female_htmls = []
            cur_father_ko = clean_ko_name(p_row.get("인명", ""))
            father_vert_text = "<br>".join(list(cur_father_ko)) if cur_father_ko else "父"

            for ch in sorted_children:
                ch_rel = str(ch['관계']).strip()
                ch_gid = str(ch['고유ID']).strip()
                ch_label = format_tree_node_label(ch)
                n_cls = get_tree_node_class(ch, is_self=False, is_direct=False, is_ancestor_row=False)

                node_html = f'<a href="?tree_gid={ch_gid}&init_gid={root_init_gid}" target="_self" class="{n_cls}">{ch_label}</a>'
                if "녀" in ch_rel:
                    step_female_htmls.append(node_html)
                    if children_by_fgid.get(ch_gid):
                        next_gen_rows.append(ch)
                else:
                    step_male_htmls.append(node_html)
                    next_gen_rows.append(ch)

            step_split_html = build_split_zones_html(step_male_htmls, step_female_htmls)
            cousin_step_htmls.append(
                f'<div class="tree-cousin-step">'
                f'<div class="tree-father-vert-box">{father_vert_text}</div>'
                f'<div class="tree-children-group-box">{step_split_html}</div>'
                f'</div>'
            )

        if not cousin_step_htmls:
            break

        steps_container_html = f'<div class="tree-cousin-steps-container">{"".join(cousin_step_htmls)}</div>'

        diagram_parts.append('<div class="tree-connector-line">▼</div>')
        dg_num = (tg_num + down_lvl) if tg_num else (24 + down_lvl)
        d_badge = down_labels.get(down_lvl, f"{down_lvl}대손")
        diagram_parts.append(f"""
        <div class="tree-level-row">
            <span class="tree-gen-badge">▼ {d_badge}<br>{dg_num}세</span>
            {steps_container_html}
        </div>
        """)
        current_parents = next_gen_rows

    diagram_parts.append('</div>')
    render_html("".join(diagram_parts))

    render_person_result_card(target_row)

def render_bunpado_view():
    render_html("""
    <div class="result-card">
        <div class="result-card-header">
            <div class="result-name-big">
                <span>🌳 광주안씨 완귀공파 분파도 (24세 ~ 30세)</span>
            </div>
            <div class="result-book-loc">시직공파보<br>1권 수록</div>
        </div>
        <div style="font-size:9.8pt; color:#dc2626; font-weight:700; margin-bottom:8px;">
            ※ 각 선조의 함자 박스를 클릭하시면 해당 선조의 전체 계보도로 즉시 이동합니다.
        </div>
        <div class="meta-para-container" style="background:#ffffff;">
            <div class="tree-level-row">
                <span class="tree-gen-badge">24세<br>파조</span>
                <div class="tree-split-wrapper">
                    <div class="tree-left-male-zone">
                        <a href="?tree_gid=A-24-1-1-24&init_gid=A-24-1-1-24" target="_self" class="tree-node-box tree-node-direct">증_嶒</a>
                        <span style="font-size:9.5pt; font-weight:700; color:#1e3c72;">완귀공 (시강원 설서·사서)</span>
                    </div>
                </div>
            </div>
            <div class="tree-connector-line">▼</div>
            <div class="tree-level-row">
                <span class="tree-gen-badge">25세<br>1대손</span>
                <div class="tree-split-wrapper">
                    <div class="tree-left-male-zone">
                        <a href="?tree_gid=A-24-2-1-25&init_gid=A-24-2-1-25" target="_self" class="tree-node-box tree-node-direct">종경_宗慶&nbsp;&nbsp;&nbsp;계자</a>
                        <span style="font-size:9.5pt; font-weight:700; color:#1e3c72;">참봉공 (헌릉참봉)</span>
                    </div>
                </div>
            </div>
            <div class="tree-connector-line">▼</div>
            <div class="tree-level-row">
                <span class="tree-gen-badge">26세<br>2대손</span>
                <div class="tree-split-wrapper">
                    <div class="tree-left-male-zone">
                        <a href="?tree_gid=A-24-3-1-26&init_gid=A-24-3-1-26" target="_self" class="tree-node-box tree-node-direct">대해_大海</a>
                        <span style="font-size:9.5pt; font-weight:700; color:#1e3c72;">참봉공 (음참봉)</span>
                    </div>
                </div>
            </div>
            <div class="tree-connector-line">▼</div>
            <div class="tree-level-row">
                <span class="tree-gen-badge">27세<br>3대손</span>
                <div class="tree-split-wrapper">
                    <div class="tree-left-male-zone">
                        <a href="?tree_gid=A-24-4-1-27&init_gid=A-24-4-1-27" target="_self" class="tree-node-box tree-node-direct">전_琠&nbsp;&nbsp;&nbsp;계자</a>
                        <span style="font-size:9.5pt; font-weight:700; color:#1e3c72;">감찰공 (사헌부 감찰)</span>
                    </div>
                </div>
            </div>
            <div class="tree-connector-line">▼</div>
            <div class="tree-level-row">
                <span class="tree-gen-badge">28세<br>4대손</span>
                <div class="tree-split-wrapper">
                    <div class="tree-left-male-zone">
                        <a href="?tree_gid=A-24-5-1-28&init_gid=A-24-5-1-28" target="_self" class="tree-node-box tree-node-male">명한_鳴漢</a>
                        <a href="?tree_gid=A-27-5-1-28&init_gid=A-27-5-1-28" target="_self" class="tree-node-box tree-node-male">시한_蓍漢</a>
                    </div>
                </div>
            </div>
            <div class="tree-connector-line">▼</div>
            <div class="tree-level-row">
                <span class="tree-gen-badge">29세<br>5대손</span>
                <div class="tree-cousin-steps-container">
                    <div class="tree-cousin-step">
                        <div class="tree-father-vert-box">명<br>한</div>
                        <div class="tree-children-group-box">
                            <div class="tree-left-male-zone">
                                <a href="?tree_gid=B-81-1-1-29&init_gid=B-81-1-1-29" target="_self" class="tree-node-box tree-node-male">세영_世英</a>
                                <a href="?tree_gid=B-123-1-1-29&init_gid=B-123-1-1-29" target="_self" class="tree-node-box tree-node-male">세익_世翼</a>
                                <a href="?tree_gid=A-24-6-1-29&init_gid=A-24-6-1-29" target="_self" class="tree-node-box tree-node-male">세웅_世雄</a>
                            </div>
                        </div>
                    </div>
                    <div class="tree-cousin-step">
                        <div class="tree-father-vert-box">시<br>한</div>
                        <div class="tree-children-group-box">
                            <div class="tree-left-male-zone">
                                <a href="?tree_gid=B-130-1-1-29&init_gid=B-130-1-1-29" target="_self" class="tree-node-box tree-node-male">세징_世徵</a>
                                <a href="?tree_gid=B-148-1-1-29&init_gid=B-148-1-1-29" target="_self" class="tree-node-box tree-node-male">세형_世亨</a>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="tree-connector-line">▼</div>
            <div class="tree-level-row">
                <span class="tree-gen-badge">30세<br>주요지파</span>
                <div class="tree-cousin-steps-container">
                    <div class="tree-cousin-step">
                        <div class="tree-father-vert-box">세<br>영</div>
                        <div class="tree-children-group-box">
                            <div class="tree-left-male-zone">
                                <a href="?tree_gid=B-81-2-1-30&init_gid=B-81-2-1-30" target="_self" class="tree-node-box tree-node-male">후정_后靜</a>
                                <a href="?tree_gid=B-119-2-1-30&init_gid=B-119-2-1-30" target="_self" class="tree-node-box tree-node-male">인정_仁靜</a>
                                <span style="font-size:9pt; color:#334155; font-weight:700;">(상석계파 등)</span>
                            </div>
                        </div>
                    </div>
                    <div class="tree-cousin-step">
                        <div class="tree-father-vert-box">세<br>익</div>
                        <div class="tree-children-group-box">
                            <div class="tree-left-male-zone">
                                <a href="?tree_gid=B-123-2-5-30&init_gid=B-123-2-5-30" target="_self" class="tree-node-box tree-node-male">태정_泰靜</a>
                                <span style="font-size:9pt; color:#334155; font-weight:700;">(목계파 등)</span>
                            </div>
                        </div>
                    </div>
                    <div class="tree-cousin-step">
                        <div class="tree-father-vert-box">세<br>징</div>
                        <div class="tree-children-group-box">
                            <div class="tree-left-male-zone">
                                <a href="?tree_gid=B-130-2-1-30&init_gid=B-130-2-1-30" target="_self" class="tree-node-box tree-node-male">처정_處靜</a>
                                <a href="?tree_gid=B-145-2-1-30&init_gid=B-145-2-1-30" target="_self" class="tree-node-box tree-node-male">유정_逌靜</a>
                                <a href="?tree_gid=B-146-2-1-30&init_gid=B-146-2-1-30" target="_self" class="tree-node-box tree-node-male">화정_和靜</a>
                                <a href="?tree_gid=B-146-2-2-30&init_gid=B-146-2-2-30" target="_self" class="tree-node-box tree-node-male">유정_有靜</a>
                                <a href="?tree_gid=B-147-2-1-30&init_gid=B-147-2-1-30" target="_self" class="tree-node-box tree-node-male">사정_思靜</a>
                                <span style="font-size:9pt; color:#334155; font-weight:700;">(봉계파 등)</span>
                            </div>
                        </div>
                    </div>
                    <div class="tree-cousin-step">
                        <div class="tree-father-vert-box">세<br>형</div>
                        <div class="tree-children-group-box">
                            <div class="tree-left-male-zone">
                                <a href="?tree_gid=B-148-2-3-30&init_gid=B-148-2-3-30" target="_self" class="tree-node-box tree-node-male">상정_尚靜</a>
                                <a href="?tree_gid=B-149-2-1-30&init_gid=B-149-2-1-30" target="_self" class="tree-node-box tree-node-male">응정
