"""
설문조사 결과 분석 대시보드 — Streamlit 버전

실행 방법:
  pip install streamlit pandas openpyxl plotly xlsxwriter
  streamlit run survey_streamlit.py
"""

import io
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────
# 페이지 설정  ← 반드시 첫 번째 st 호출
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Survey Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# 디자인 토큰
# ─────────────────────────────────────────────
ACCENT  = "#3B82F6"
TBL_HDR = "#1E3A5F"
PALETTE = ["#3B82F6","#10B981","#F59E0B","#EF4444","#8B5CF6",
           "#06B6D4","#F97316","#EC4899","#14B8A6","#6366F1"]

# ─────────────────────────────────────────────
# CSS 주입  ← set_page_config 바로 다음, 다른 st 호출 전
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans','Noto Sans KR',sans-serif !important; }

/* 사이드바 다크 테마 */
[data-testid="stSidebar"] { background:#0F1B2D !important; border-right:1px solid #1E3050; }
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div { color:#CBD5E1; }
[data-testid="stSidebar"] label { font-size:11px !important; font-weight:600 !important; color:#64748B !important; text-transform:uppercase; letter-spacing:.07em; }
[data-testid="stSidebar"] textarea { background:#162030 !important; border:1px solid #243650 !important; color:#CBD5E1 !important; border-radius:6px !important; font-size:12px !important; }
[data-testid="stSidebar"] [data-baseweb="select"] > div { background:#162030 !important; border-color:#243650 !important; border-radius:6px !important; }
[data-testid="stSidebar"] input[type="number"] { background:#162030!important; border-color:#243650!important; color:#CBD5E1!important; border-radius:6px!important; }
[data-testid="stSidebar"] hr { border-color:#1E3050 !important; margin:6px 0 !important; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color:#F1F5F9 !important; }

/* 메인 컨테이너 */
.main .block-container { padding:24px 32px !important; max-width:100% !important; }

/* 탭 스타일 */
.stTabs [data-baseweb="tab-list"] { border-bottom:2px solid #E5E7EB; gap:2px; }
.stTabs [data-baseweb="tab"] { font-size:13px !important; font-weight:600 !important; padding:10px 18px !important; border-radius:0 !important; border:none !important; color:#6B7280 !important; background:transparent !important; }
.stTabs [aria-selected="true"] { color:#3B82F6 !important; border-bottom:2px solid #3B82F6 !important; }

/* 다운로드 버튼 */
.stDownloadButton > button { border-radius:8px !important; font-size:12px !important; font-weight:600 !important; padding:7px 16px !important; }

/* dataframe 테이블 */
[data-testid="stDataFrame"] { border-radius:0 0 10px 10px; overflow:hidden; border:1px solid #BFDBFE; border-top:none; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 유틸 함수
# ─────────────────────────────────────────────
def sorted_levels(series):
    vals = list(series.dropna().astype(str).unique())
    try:
        return sorted(vals, key=lambda x: float(x))
    except Exception:
        return sorted(vals)

def apply_rename(df, rename_text):
    if not rename_text:
        return df
    df = df.copy()
    for line in rename_text.strip().splitlines():
        if "=" in line:
            old, new = [s.strip() for s in line.split("=", 1)]
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)
    return df

def apply_labels(df, label_text):
    if not label_text:
        return df
    df = df.copy()
    for line in label_text.strip().splitlines():
        if ":" not in line:
            continue
        var_name, mappings = [s.strip() for s in line.split(":", 1)]
        if var_name not in df.columns:
            continue
        df[var_name] = df[var_name].astype(str)
        for pair in mappings.split(","):
            if "=" in pair:
                old_v, new_v = [s.strip() for s in pair.split("=", 1)]
                df[var_name] = df[var_name].replace(old_v, new_v)
    return df

def read_uploaded_file(uploaded):
    """xlsx / xls / csv 자동 판별 읽기, CSV 인코딩 자동 감지"""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        raw = uploaded.read()
        for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=enc)
            except Exception:
                continue
        raise ValueError("CSV 인코딩을 인식할 수 없습니다.")
    else:
        return pd.read_excel(uploaded)

def fmt(val, mode, dec):
    return f"{val:.{dec}f}%" if mode == "pct" else f"{val:.{dec}f}"

def apply_combine(df, combine_text):
    """
    변수 결합 생성. 형식 (줄바꿈으로 여러 개):
      새변수명 = 변수A * 변수B
      성by연령 = 성별 * 연령대
    결합 값은 "값A_값B" 형태로 생성
    """
    if not combine_text:
        return df
    df = df.copy()
    for line in combine_text.strip().splitlines():
        line = line.strip()
        if not line or "=" not in line or "*" not in line:
            continue
        new_var, expr = [s.strip() for s in line.split("=", 1)]
        parts = [s.strip() for s in expr.split("*")]
        if not new_var or len(parts) < 2:
            continue
        missing = [p for p in parts if p not in df.columns]
        if missing:
            continue
        df[new_var] = df[parts[0]].astype(str)
        for p in parts[1:]:
            df[new_var] = df[new_var] + "_" + df[p].astype(str)
        # 결측이 포함된 행은 결합 결과도 결측 처리
        for p in parts:
            mask = df[p].isna() | (df[p].astype(str).str.strip() == "")
            df.loc[mask, new_var] = None
    return df

def build_crosstab(df, t_var, c_vars, mode, dec, pct_base="행 기준"):
    """
    pct_base:
      행 기준  — 각 집단(행) 내 유효 응답 합계를 분모
      열 기준  — 같은 c_var 블록 안에서 각 열(t_var 값)의 합 = 100%
      전체 기준 — 전체 유효 응답자 합계를 분모
    결측(NaN, 빈 문자열)은 모든 기준에서 분모·분자 모두 제외
    """
    if t_var not in df.columns:
        return None

    # t_var 결측 제외
    valid_mask = df[t_var].notna() & (df[t_var].astype(str).str.strip() != "")
    df_valid   = df[valid_mask].copy()

    t_levels    = sorted_levels(df_valid[t_var])
    fill        = fmt(0, mode, dec)
    grand_total = df_valid["wt"].sum()

    def calc(w, denom):
        """공통 계산: w/denom*100 or w"""
        if mode != "pct":
            return fmt(w, mode, dec)
        return fmt(w / denom * 100 if denom > 0 else 0, mode, dec)

    def make_row(label, grp, col_totals_local, grand_local):
        """
        grp           : 해당 행(집단 or 전체)의 유효 응답 df
        col_totals_local : 열 기준용 — 같은 블록(c_var 전체) 내 각 tl의 가중치 합
        grand_local   : 전체 기준용 분모
        """
        row_total = grp["wt"].sum()
        row = {"구분": label, "사례수(N)": f"{row_total:.{dec}f}"}
        for tl in t_levels:
            w = grp[grp[t_var].astype(str) == str(tl)]["wt"].sum()
            if pct_base == "행 기준":
                row[tl] = calc(w, row_total)
            elif pct_base == "열 기준":
                row[tl] = calc(w, col_totals_local.get(tl, 0))
            else:  # 전체 기준
                row[tl] = calc(w, grand_local)
        return row

    # ── 전체 행
    # 열 기준 전체: df_valid 전체 기준 col_totals
    col_totals_all = {
        tl: df_valid[df_valid[t_var].astype(str) == str(tl)]["wt"].sum()
        for tl in t_levels
    }
    all_rows = [make_row("전체", df_valid, col_totals_all, grand_total)]

    # ── c_var 블록별 행
    for c_var in c_vars:
        if c_var not in df_valid.columns:
            continue
        # c_var 결측 제외
        block = df_valid[
            df_valid[c_var].notna() & (df_valid[c_var].astype(str).str.strip() != "")
        ]
        # 열 기준: 이 블록(c_var 유효 전체) 안에서의 열 합계를 분모로 사용
        # → 블록 내 모든 집단 행을 합치면 각 열이 100%가 됨
        col_totals_block = {
            tl: block[block[t_var].astype(str) == str(tl)]["wt"].sum()
            for tl in t_levels
        }
        grand_block = block["wt"].sum()

        for lvl in sorted_levels(block[c_var]):
            grp = block[block[c_var].astype(str) == str(lvl)]
            label = f"[{c_var}] {lvl}"
            all_rows.append(make_row(label, grp, col_totals_block, grand_block))

    return pd.DataFrame(all_rows, columns=["구분", "사례수(N)"] + t_levels).fillna(fill)

def style_crosstab(df):
    """전체 행 강조 + 헤더 스타일"""
    def highlight_total(row):
        if row["구분"] == "전체":
            return [f"background-color:#EFF6FF; font-weight:700; color:{TBL_HDR}"] * len(row)
        return [""] * len(row)
    return (
        df.style
        .apply(highlight_total, axis=1)
        .set_table_styles([
            {"selector": "thead th", "props": [
                ("background-color", TBL_HDR), ("color", "white"),
                ("font-size", "12px"), ("font-weight", "600"),
                ("text-align", "center"), ("padding", "10px 8px"),
            ]},
            {"selector": "td", "props": [
                ("text-align", "center"), ("font-size", "12.5px"),
                ("padding", "7px 8px"), ("border-color", "#F3F4F6"),
            ]},
            {"selector": "td:first-child", "props": [
                ("text-align", "left"), ("padding-left", "14px"), ("font-weight", "500"),
            ]},
            {"selector": "tr:nth-child(even) td", "props": [("background-color", "#FAFAFA")]},
            {"selector": "tr:hover td",           "props": [("background-color", "#EFF6FF")]},
        ])
        .hide(axis="index")
    )

def to_excel_bytes(all_tables):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt   = wb.add_format({"bold": True, "bg_color": TBL_HDR, "font_color": "white",
                                    "align": "center", "border": 1, "font_size": 11})
        total_fmt = wb.add_format({"bold": True, "bg_color": "#EFF6FF", "font_color": TBL_HDR,
                                    "align": "center", "border": 1, "font_size": 11})
        for t_var, df in all_tables.items():
            sheet = str(t_var)[:31]
            df.to_excel(writer, index=False, sheet_name=sheet)
            ws = writer.sheets[sheet]
            for c, col in enumerate(df.columns):
                ws.write(0, c, col, hdr_fmt)
            for r, row in enumerate(df.itertuples(index=False), 1):
                if str(row[0]) == "전체":
                    for c, val in enumerate(row):
                        ws.write(r, c, val, total_fmt)
            ws.set_column(0, 0, 30)
            ws.set_column(1, len(df.columns) - 1, 12)
            ws.freeze_panes(1, 0)
    buf.seek(0)
    return buf.read()

def to_csv_bytes(all_tables):
    parts = []
    for t_var, df in all_tables.items():
        sep = pd.DataFrame([["▶ " + t_var] + [""] * (len(df.columns) - 1)], columns=df.columns)
        parts.extend([sep, df])
    return pd.concat(parts, ignore_index=True).to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

def build_chart(df, t_var, p_group, dec):
    t_levels = sorted_levels(df[t_var])
    base = dict(
        paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="DM Sans, Noto Sans KR, sans-serif", size=12, color="#111827"),
        margin=dict(l=60, r=30, t=55, b=60),
        hoverlabel=dict(bgcolor="white", bordercolor="#E5E7EB", font=dict(size=12)),
    )

    # 전체 단순 막대
    if p_group == "전체":
        grp = df[df[t_var].notna()].groupby(df[t_var].astype(str))["wt"].sum().reset_index()
        grp.columns = [t_var, "w_n"]
        grp["pct"] = grp["w_n"] / grp["w_n"].sum() * 100
        grp = grp.set_index(t_var).reindex(t_levels).reset_index()
        fig = go.Figure()
        for i, row in grp.iterrows():
            fig.add_trace(go.Bar(
                x=[row[t_var]], y=[row["pct"]],
                marker=dict(color=PALETTE[i % len(PALETTE)], line=dict(width=0)),
                text=[f"<b>{row['pct']:.{dec}f}%</b>"], textposition="outside",
                hovertemplate=f"<b>{row[t_var]}</b><br>{row['pct']:.{dec}f}%<extra></extra>",
                showlegend=False, width=0.5,
            ))
        fig.update_layout(
            **base,
            title=dict(text=f"<b>{t_var}</b>  전체 응답 비율",
                       font=dict(size=14, color=TBL_HDR), x=0, xanchor="left"),
            xaxis=dict(title=None, showgrid=False, zeroline=False),
            yaxis=dict(title="비율 (%)", showgrid=True, gridcolor="#F3F4F6", zeroline=False,
                       range=[0, (grp["pct"].max() if not grp.empty else 100) * 1.2]),
            bargap=0.4,
        )
        return fig

    # 교차 누적 가로 막대
    if p_group not in df.columns:
        return None
    p_levels = sorted_levels(df[p_group])
    sub = df[df[t_var].notna() & df[p_group].notna()].copy()
    sub[t_var]   = sub[t_var].astype(str)
    sub[p_group] = sub[p_group].astype(str)
    grp = sub.groupby([p_group, t_var])["wt"].sum().reset_index()
    grp.columns = [p_group, t_var, "w_n"]
    totals = grp.groupby(p_group)["w_n"].sum().reset_index().rename(columns={"w_n": "total"})
    grp = grp.merge(totals, on=p_group)
    grp["pct"] = grp["w_n"] / grp["total"] * 100
    fig = go.Figure()
    for i, tl in enumerate(t_levels):
        sub_tl = grp[grp[t_var] == str(tl)].set_index(p_group).reindex(p_levels)
        vals = sub_tl["pct"].fillna(0).values
        fig.add_trace(go.Bar(
            y=p_levels, x=vals, name=str(tl), orientation="h",
            marker=dict(color=PALETTE[i % len(PALETTE)], line=dict(width=0)),
            text=[f"<b>{v:.{dec}f}%</b>" if v > 4 else "" for v in vals],
            textposition="inside", insidetextanchor="middle",
            hovertemplate=f"<b>{tl}</b><br>%{{y}}: %{{x:.{dec}f}}%<extra></extra>",
        ))
    fig.update_layout(
        **base, barmode="stack",
        title=dict(text=f"<b>{p_group}</b> 특성별  <b>{t_var}</b>  응답 분포",
                   font=dict(size=14, color=TBL_HDR), x=0, xanchor="left"),
        xaxis=dict(title="비율 (%)", range=[0, 100], showgrid=True,
                   gridcolor="#F3F4F6", zeroline=False, ticksuffix="%"),
        yaxis=dict(title=None, showgrid=False, zeroline=False),
        legend=dict(title=dict(text=t_var), orientation="h",
                    yanchor="bottom", y=-0.22, xanchor="left", x=0,
                    font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
        bargap=0.25,
    )
    return fig


# ─────────────────────────────────────────────
# 사이드바 — st.* 위젯으로만 구성 (HTML 클래스 의존 제거)
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Survey Analytics")
    st.caption("결과 분석 대시보드 · Cross-tab & Visualization")
    st.divider()

    # 1. 파일 업로드
    st.markdown("**☁ 데이터 업로드**")
    uploaded = st.file_uploader("Excel / CSV 파일 선택",
                                type=["xlsx", "xls", "csv"],
                                label_visibility="collapsed")
    st.divider()

    # 2. 변수명 변경
    st.markdown("**✏ 변수명 변경**")
    rename_text = st.text_area(
        "기존변수=새변수 (줄바꿈)",
        placeholder="SQ1=지역\nSQ2=연령대",
        height=88,
        label_visibility="collapsed",
    )
    st.divider()

    # 3. 라벨링
    st.markdown("**🏷 값 라벨링**")
    label_text = st.text_area(
        "변수명: 1=값A, 2=값B",
        placeholder="성별: 1=남자, 2=여자\n지역: 1=서울, 2=경기",
        height=110,
        label_visibility="collapsed",
    )
    st.divider()

    # 4. 변수 결합
    st.markdown("**🔗 변수 결합**")
    show_combine = st.toggle("변수 결합 사용", value=False)
    if show_combine:
        combine_text = st.text_area(
            "새변수명 = 변수A * 변수B",
            placeholder="성by연령 = 성별 * 연령대\n지역by성별 = 지역 * 성별",
            height=100,
            label_visibility="collapsed",
            help="* 로 결합할 변수를 구분합니다. 결합된 변수는 분석 변수 목록에 자동 추가됩니다.",
        )
    else:
        combine_text = ""
    st.divider()

    # 5. 가중치 & 필터 (파일 로드 후 활성화)
    st.markdown("**⚖ 가중치 & 필터**")
    weight_var  = "(가중치 없음)"
    filter_var  = "(선택 안 함)"
    filter_vals = []

    if uploaded:
        _raw = read_uploaded_file(uploaded)
        _raw = apply_rename(_raw, rename_text)
        _raw = apply_labels(_raw, label_text)
        _raw = apply_combine(_raw, combine_text)
        _cols = _raw.columns.tolist()

        weight_var = st.selectbox("가중치 변수", ["(가중치 없음)"] + _cols)
        filter_var = st.selectbox("필터 변수",   ["(선택 안 함)"]  + _cols)
        if filter_var != "(선택 안 함)":
            _choices    = sorted_levels(_raw[filter_var])
            filter_vals = st.multiselect(
                f"'{filter_var}' 포함 값", _choices, default=_choices)
    else:
        st.caption("파일을 업로드하면\n변수 목록이 표시됩니다.")

    st.divider()

    # 6. 출력 설정
    st.markdown("**⚙ 출력 설정**")
    display_mode = st.radio("표시 방식", ["비율 (%)", "사례수 (N)"], horizontal=True)
    mode = "pct" if "비율" in display_mode else "count"
    if mode == "pct":
        pct_base = st.radio(
            "비율 기준",
            ["행 기준", "열 기준", "전체 기준"],
            help=(
                "행 기준: 각 집단(행) 합계 = 100%  ← 기본\n"
                "열 기준: 각 응답(열) 합계 = 100%\n"
                "전체 기준: 전체 응답자 합계 = 100%"
            ),
        )
    else:
        pct_base = "행 기준"
    dec = int(st.number_input("소수점 자리수", min_value=0, max_value=5, value=1, step=1))


# ─────────────────────────────────────────────
# 메인 영역
# ─────────────────────────────────────────────

# 페이지 제목 (st.* 로 렌더링 — HTML 클래스 불필요)
st.markdown("### 분석 워크스페이스")
st.caption("교차분석표 & 시각화")

# ── 파일 미업로드
if not uploaded:
    st.info("👈 좌측 사이드바에서 Excel 파일을 업로드하면 분석을 시작할 수 있습니다.")
    st.stop()

# ── 데이터 전처리
df = _raw.copy()
if filter_var != "(선택 안 함)" and filter_vals and filter_var in df.columns:
    df = df[df[filter_var].astype(str).isin([str(v) for v in filter_vals])]
df["wt"] = (
    pd.to_numeric(df[weight_var], errors="coerce").fillna(0)
    if weight_var != "(가중치 없음)" and weight_var in df.columns
    else 1.0
)
# wt 컬럼 추가 이전 원본 컬럼만 사용 (wt 제외)
cols = [c for c in _raw.columns.tolist() if c != "wt"]

# ── 변수 설정 패널
with st.container(border=True):
    st.caption("🔲 분석 변수 설정")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("🔵 **COL** — 가로(Column) 변수 · 주요 분석 문항")
        target_vars = st.multiselect(
            "target", cols, placeholder="분석 문항 선택 (다중)...",
            label_visibility="collapsed")
    with c2:
        st.markdown("🟣 **ROW** — 세로(Row) 변수 · 집단 특성 문항")
        cross_vars = st.multiselect(
            "cross", cols, placeholder="집단 변수 선택 (다중)...",
            label_visibility="collapsed")

# ── 선택 미완료
if not target_vars or not cross_vars:
    st.info("📊 가로(COL) 변수와 세로(ROW) 변수를 각각 하나 이상 선택하면 교차 분석표가 생성됩니다.")
    st.stop()

# ── 교차표 생성
all_tables = {}
for t_var in target_vars:
    result = build_crosstab(df, t_var, cross_vars, mode, dec, pct_base)
    if result is not None:
        all_tables[t_var] = result

# ─────────────────────────────────────────────
# 탭
# ─────────────────────────────────────────────
tab_table, tab_chart = st.tabs(["📋  데이터 테이블", "📈  시각화"])

# ── 탭 1: 교차 분석표
with tab_table:
    col_xl, col_csv, col_note = st.columns([1, 1, 6])
    with col_xl:
        st.download_button(
            "📥 Excel",
            data=to_excel_bytes(all_tables),
            file_name="교차분석표.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_csv:
        st.download_button(
            "📋 CSV",
            data=to_csv_bytes(all_tables),
            file_name="교차분석표.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col_note:
        st.caption("문항별 시트/구분행으로 내보내기")

    for i, (t_var, result) in enumerate(all_tables.items(), 1):
        # 카드 헤더 — inline style로 렌더링 (외부 클래스 불필요)
        st.markdown(f"""
        <div style="
            background:linear-gradient(135deg,#EFF6FF 0%,#F0F9FF 100%);
            border:1px solid #BFDBFE;
            border-radius:10px 10px 0 0;
            padding:12px 18px;
            display:flex;
            align-items:center;
            gap:10px;
            margin-top:20px;
        ">
            <span style="
                background:#3B82F6;color:white;font-size:11px;font-weight:800;
                width:24px;height:24px;border-radius:6px;
                display:inline-flex;align-items:center;justify-content:center;
            ">{i}</span>
            <span style="font-size:14px;font-weight:700;color:#1E3A5F;">{t_var}</span>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(style_crosstab(result), use_container_width=True, hide_index=True)

# ── 탭 2: 시각화
with tab_chart:
    vc1, vc2, vc3 = st.columns([2, 2, 3])
    with vc1:
        plot_tvar = st.selectbox("시각화 문항", target_vars)
    with vc2:
        plot_group = st.selectbox("그룹 기준", ["전체"] + cross_vars)
    with vc3:
        st.caption(" ")
        st.caption("ℹ 1번 응답이 왼쪽/위부터 배치됩니다.")

    if plot_tvar:
        fig = build_chart(df, plot_tvar, plot_group, dec)
        if fig:
            st.plotly_chart(
                fig, use_container_width=True,
                config={"displaylogo": False,
                        "modeBarButtonsToRemove": ["select2d", "lasso2d"]},
            )
        else:
            st.warning("선택한 그룹 변수가 데이터에 없습니다.")