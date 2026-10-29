# 2. 메인 타이틀 (중앙 정렬 및 글자 크기 조절)
title_html = """
<div style="text-align: center; margin-top: 10px; margin-bottom: 20px;">
    <span style="font-size: 32px; font-weight: 700; color: #1a237e; letter-spacing: -0.5px;">
        📜 광주안씨 완귀공파 문중
    </span><br>
    <span style="font-size: 26px; font-weight: 600; color: #212121; letter-spacing: -0.5px;">
        족보 통합 검색 시스템
    </span>
    <p style="font-size: 15px; color: #666666; margin-top: 10px;">
        전산화된 족보 데이터를 언제 어디서나 편리하게 다각도로 조회하세요.
    </p>
</div>
"""
st.markdown(title_html, unsafe_allow_html=True)
