import streamlit as st

st.set_page_config(page_title="AI 工具箱", layout="wide")

# 复用你之前的精美 CSS
st.markdown("""
<style>
    [data-testid="stHeader"], footer { visibility: hidden !important; }
    .page-title { text-align: center; margin-bottom: 30px; }
    .card-grid {
        display: flex;
        justify-content: center;
        gap: 18px;
        flex-wrap: wrap;
        max-width: 900px;
        margin: 0 auto;
    }
    .card {
        background: #ffffff;
        border-radius: 14px;
        padding: 20px 15px;
        text-align: center;
        text-decoration: none !important;
        color: #333 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        transition: all 0.25s ease;
        width: 180px;
    }
    .card:hover { transform: translateY(-5px); box-shadow: 0 10px 24px rgba(0,0,0,0.1); }
    .icon { font-size: 35px; margin-bottom: 10px; }
    .card-h3 { font-size: 16px; font-weight: 600; margin-bottom: 5px; }
    .card-p { font-size: 12px; color: #888; margin: 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="page-title">
    <h1>🧰 AI 工具箱</h1>
    <p>内置多页架构 · 极速响应</p>
</div>


<div class="card-grid">

<a href="/分表工具" target="_self" class="card">
    <div class="icon">📊</div>
    <div class="card-h3">Excel 分表工具</div>
    <p class="card-p">上传 Excel，按字段拆分</p>
</a>

<div class="card disabled-card">
    <div class="icon">🛠️</div>
    <div class="card-h3">更多工具</div>
    <p class="card-p">即将上线</p>
</div>

<div class="card disabled-card">
    <div class="icon">ℹ️</div>
    <div class="card-h3">关于</div>
    <p class="card-p">简洁实用的工具集合</p>
</div>

</div>
""", unsafe_allow_html=True)