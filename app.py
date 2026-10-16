# -*- coding: utf-8 -*-
"""지점 원두 우체국 접수 — Streamlit 앱"""
import os, datetime
from zoneinfo import ZoneInfo
import pandas as pd
import streamlit as st
import core

BASE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ADDR = os.path.join(BASE, "지점우체국택배접수주소.xlsx")

st.set_page_config(page_title="지점 원두 우체국 접수", page_icon="📦", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 2rem; max-width: 1200px;}
h1 {color:#8B3A2A;}
div[data-testid="stMetricValue"] {color:#8B3A2A; font-weight:700;}
div[data-testid="stMetric"] {background:#FBF8F5; border:1px solid #D6CEC8; border-radius:8px; padding:12px 16px;}
</style>
""", unsafe_allow_html=True)


# ── 비밀번호 (Streamlit Cloud > Settings > Secrets 에 APP_PASSWORD 설정 시에만 작동) ──
def check_password():
    try:
        pw = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        pw = ""
    if not pw or st.session_state.get("auth_ok"):
        return True
    st.title("지점 원두 우체국 접수")
    entered = st.text_input("비밀번호", type="password")
    if entered:
        if entered == pw:
            st.session_state["auth_ok"] = True
            st.rerun()
        st.error("비밀번호가 맞지 않습니다.")
    return False


if not check_password():
    st.stop()

# ── 사이드바: 사용법 / 주소파일 ──
with st.sidebar:
    st.subheader("사용 방법")
    st.markdown(
        "1. 오늘 **지점원두발주서**(.xlsx)를 올립니다\n"
        "2. 빨간 경고(주소 미등록 지점)가 있는지 확인합니다\n"
        "3. **접수파일 다운로드** → 우체국에 업로드\n"
    )
    st.divider()
    st.subheader("박스 기준")
    st.caption(
        "• 어센틱에스프레소블렌드 1kg: 15kg 기준 별도 박스, 박스마다 고르게 분배\n\n"
        "• 그 외 중량 표기 원두: 지점별 합산 후 15kg 기준\n\n"
        "• 강릉(본점·경포·임당·사천·신라모노) 제외"
    )
    st.divider()
    st.subheader("지점 주소파일")
    # 주소록 우선순위: ① 여기서 올린 파일 → ② Secrets의 [[addresses]] → ③ 폴더의 주소파일
    try:
        secret_addr = [dict(a) for a in st.secrets.get("addresses", [])]
    except Exception:
        secret_addr = []
    addr_up = st.file_uploader("주소파일 교체 (이번 작업에만 적용)", type=["xlsx"], key="addr")
    if addr_up:
        st.caption("✅ 올린 주소파일 사용 중")
    elif secret_addr:
        st.caption(f"✅ 등록된 주소록 사용 중 ({len(secret_addr)}곳)")
    elif os.path.exists(DEFAULT_ADDR):
        st.caption("✅ 기본 주소파일 사용 중")
    else:
        st.error("등록된 주소록이 없습니다. 주소파일을 올려주세요.")

if addr_up:
    addr_src = addr_up
elif secret_addr:
    addr_src = secret_addr
elif os.path.exists(DEFAULT_ADDR):
    addr_src = DEFAULT_ADDR
else:
    addr_src = None

# ── 본문 ──
st.title("📦 지점 원두 우체국 접수")
st.caption("지점 원두 발주서를 올리면 지점별·박스별 우체국 접수파일을 만들어 줍니다.")

order_up = st.file_uploader("지점원두발주서 (.xlsx)", type=["xlsx"], key="order")
if not order_up:
    st.info("발주서 파일을 위에 끌어다 놓으세요.")
    st.stop()
if addr_src is None:
    st.stop()

try:
    res = core.process(order_up, addr_src)
except Exception as e:
    st.error(f"파일을 처리하지 못했습니다: {e}")
    st.stop()

summary = pd.DataFrame(res["summary"], columns=[
    "발주서 지점", "수취인명", "어센틱1kg(kg)", "어센틱 박스", "기타원두(kg)", "기타 박스", "총 박스", "비고"])
target = summary[summary["총 박스"] > 0]

# 날짜 확인
if res["date"]:
    today = datetime.datetime.now(ZoneInfo("Asia/Seoul")).date()
    msg = f"발주일자: **{res['date']:%Y-%m-%d}**"
    if res["date"] != today:
        st.warning(msg + f"  ← 오늘({today:%m/%d})과 다릅니다. 파일을 다시 확인하세요.")
    else:
        st.success(msg)

# 주소 미등록 경고
if res["missing"]:
    lines = "\n".join(f"- **{b}** : 어센틱 {a:g}kg / 기타 {o:g}kg" for b, a, o in res["missing"])
    st.error("**주소 미등록 지점 — 접수파일에서 빠졌습니다. 별도 접수하세요.**\n\n" + lines)

c1, c2, c3, c4 = st.columns(4)
c1.metric("접수 지점", f"{len(target)}곳")
c2.metric("총 박스", f"{len(res['upload'])}박스")
c3.metric("어센틱 1kg", f"{target['어센틱1kg(kg)'].sum():g}kg")
c4.metric("기타 원두", f"{target['기타원두(kg)'].sum():g}kg")

fname = f"{res['mmdd']}_지점원두우체국접수파일.xlsx"
st.download_button(
    "⬇️ 접수파일 다운로드", data=core.build_workbook(res, order_up.name),
    file_name=fname, type="primary", use_container_width=True,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

tab1, tab2 = st.tabs(["지점별 박스", "송장 미리보기"])
with tab1:
    show_all = st.toggle("제외·발주없음 지점도 보기", value=False)
    view = summary if show_all else target
    st.dataframe(view, hide_index=True, use_container_width=True,
                 column_config={"어센틱1kg(kg)": st.column_config.NumberColumn(format="%g"),
                                "기타원두(kg)": st.column_config.NumberColumn(format="%g")})
with tab2:
    st.dataframe(pd.DataFrame(res["upload"], columns=core.HEADERS),
                 hide_index=True, use_container_width=True, height=500)

if res["skipped"]:
    st.caption("※ 중량 표기가 없어 계산에서 제외된 품목: " + ", ".join(res["skipped"]))
