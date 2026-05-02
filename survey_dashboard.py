"""
설문조사 결과 분석 대시보드
R Shiny(SurveyTab.R) → Python Dash 변환 (리디자인)

실행 방법:
  pip install dash dash-bootstrap-components pandas openpyxl plotly xlsxwriter
  python survey_dashboard.py
"""

import io
import base64

import dash
from dash import dcc, html, dash_table, Input, Output, State, no_update
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go

# ─────────────────────────────────────────────
# 앱 초기화
# ─────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        "https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap",
        "https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css",
    ],
    suppress_callback_exceptions=True,
)
app.title = "Survey Analytics"

# ── 디자인 토큰
NAV_BG      = "#0F1B2D"
NAV_BORDER  = "#1E3050"
ACCENT      = "#3B82F6"
ACCENT_DARK = "#2563EB"
SURFACE     = "#F8F9FC"
CARD_BG     = "#FFFFFF"
TEXT_MAIN   = "#111827"
TEXT_SUB    = "#6B7280"
BORDER      = "#E5E7EB"
TBL_HDR     = "#1E3A5F"
FONT_BODY   = "'DM Sans', 'Noto Sans KR', sans-serif"

GLOBAL_CSS = f"""
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:{FONT_BODY};background:{SURFACE};color:{TEXT_MAIN};font-size:14px;line-height:1.6}}
::-webkit-scrollbar{{width:6px;height:6px}}
::-webkit-scrollbar-track{{background:#f1f1f1}}
::-webkit-scrollbar-thumb{{background:#c1c9d4;border-radius:3px}}

#sidebar{{width:300px;min-width:300px;background:{NAV_BG};min-height:100vh;display:flex;flex-direction:column;border-right:1px solid {NAV_BORDER}}}
#sidebar-logo{{padding:24px 20px 20px;border-bottom:1px solid {NAV_BORDER}}}
.logo-mark{{font-size:10px;font-weight:700;letter-spacing:.15em;color:{ACCENT};text-transform:uppercase;margin-bottom:4px}}
.logo-title{{font-size:17px;font-weight:700;color:#F1F5F9;line-height:1.3}}
.logo-sub{{font-size:11px;color:#64748B;margin-top:2px;font-weight:400}}

.sb-section{{padding:14px 20px 12px;border-bottom:1px solid {NAV_BORDER}}}
.sb-section:last-child{{border-bottom:none}}
.sb-section-title{{font-size:10px;font-weight:700;letter-spacing:.12em;color:#475569;text-transform:uppercase;margin-bottom:10px;display:flex;align-items:center;gap:6px}}
.sb-section-title i{{font-size:12px;color:{ACCENT}}}
.sb-label{{font-size:10px;font-weight:600;color:#64748B;margin-bottom:4px;margin-top:8px;text-transform:uppercase;letter-spacing:.06em}}
.sb-label:first-child{{margin-top:0}}

.sb-textarea{{background:#162030!important;border:1px solid #243650!important;color:#CBD5E1!important;border-radius:6px!important;font-size:12px!important;resize:vertical;padding:8px 10px!important;width:100%;transition:border-color .2s;outline:none!important}}
.sb-textarea::placeholder{{color:#3D5470!important}}
.sb-textarea:focus{{border-color:{ACCENT}!important;box-shadow:0 0 0 2px rgba(59,130,246,.15)!important;background:#1a2840!important}}

.upload-btn{{width:100%;background:linear-gradient(135deg,{ACCENT} 0%,{ACCENT_DARK} 100%)!important;border:none!important;border-radius:8px!important;padding:10px 14px!important;font-size:12px!important;font-weight:600!important;color:white!important;letter-spacing:.02em;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:center;gap:7px}}
.upload-btn:hover{{opacity:.88;transform:translateY(-1px);box-shadow:0 4px 14px rgba(59,130,246,.4)!important}}
.file-badge{{background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.25);border-radius:6px;padding:5px 10px;font-size:11px;color:#93C5FD;margin-top:8px;display:flex;align-items:center;gap:5px;word-break:break-all}}

.sb-radio .form-check{{display:inline-flex;align-items:center;gap:4px;margin-right:12px}}
.sb-radio .form-check-input{{background-color:#162030;border-color:#243650;cursor:pointer}}
.sb-radio .form-check-input:checked{{background-color:{ACCENT};border-color:{ACCENT}}}
.sb-radio .form-check-label{{color:#CBD5E1;font-size:12px;cursor:pointer}}
.sb-number{{background:#162030!important;border:1px solid #243650!important;color:#CBD5E1!important;border-radius:6px!important;font-size:12px!important;padding:5px 10px!important;width:80px!important;outline:none!important}}
.sb-number:focus{{border-color:{ACCENT}!important;box-shadow:0 0 0 2px rgba(59,130,246,.15)!important}}

#main-content{{flex:1;display:flex;flex-direction:column;min-height:100vh;background:{SURFACE}}}
#topbar{{background:{CARD_BG};border-bottom:1px solid {BORDER};padding:0 28px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 1px 4px rgba(0,0,0,.04)}}
.page-title{{font-size:15px;font-weight:700;color:{TEXT_MAIN};letter-spacing:-.01em}}
.page-sub{{font-size:12px;color:{TEXT_SUB};margin-left:10px;font-weight:400}}
#content-area{{padding:24px 28px;flex:1}}

.var-card{{background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px 24px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.var-card-title{{font-size:11px;font-weight:700;color:{TEXT_SUB};text-transform:uppercase;letter-spacing:.08em;margin-bottom:14px;display:flex;align-items:center;gap:7px}}
.var-card-title i{{color:{ACCENT};font-size:14px}}
.var-label{{font-size:12px;font-weight:600;color:{TEXT_SUB};margin-bottom:5px;display:flex;align-items:center;gap:6px}}
.badge-col{{background:#DBEAFE;color:{ACCENT_DARK};font-size:10px;font-weight:700;padding:1px 6px;border-radius:4px;letter-spacing:.04em}}
.badge-row{{background:#F3E8FF;color:#7C3AED;font-size:10px;font-weight:700;padding:1px 6px;border-radius:4px;letter-spacing:.04em}}

.nav-tabs{{border-bottom:2px solid {BORDER}!important;gap:2px;padding-bottom:0}}
.nav-tabs .nav-link{{font-size:13px;font-weight:600;color:{TEXT_SUB}!important;border:none!important;border-bottom:2px solid transparent!important;border-radius:0!important;padding:10px 18px!important;margin-bottom:-2px;transition:all .15s;background:transparent!important;display:flex;align-items:center;gap:6px}}
.nav-tabs .nav-link:hover{{color:{TEXT_MAIN}!important;border-bottom-color:#D1D5DB!important}}
.nav-tabs .nav-link.active{{color:{ACCENT}!important;border-bottom:2px solid {ACCENT}!important;background:transparent!important}}
.tab-content{{padding-top:20px}}

.dl-btn{{border-radius:8px!important;font-size:12px!important;font-weight:600!important;padding:7px 14px!important;display:inline-flex;align-items:center;gap:6px;transition:all .2s!important;border:none!important;cursor:pointer}}
.dl-btn-excel{{background:#059669!important;color:white!important}}
.dl-btn-excel:hover{{background:#047857!important;transform:translateY(-1px);box-shadow:0 4px 10px rgba(5,150,105,.3)!important}}
.dl-btn-csv{{background:{ACCENT}!important;color:white!important}}
.dl-btn-csv:hover{{background:{ACCENT_DARK}!important;transform:translateY(-1px);box-shadow:0 4px 10px rgba(59,130,246,.3)!important}}

.result-card{{background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04);margin-bottom:20px}}
.result-card-header{{padding:14px 20px;background:linear-gradient(135deg,#EFF6FF 0%,#F0F9FF 100%);border-bottom:1px solid #DBEAFE;display:flex;align-items:center;gap:10px}}
.q-index{{background:{ACCENT};color:white;font-size:10px;font-weight:800;width:22px;height:22px;border-radius:6px;display:flex;align-items:center;justify-content:center;flex-shrink:0}}
.q-title{{font-size:13px;font-weight:700;color:{TBL_HDR}}}
.result-card-body{{padding:16px}}

.empty-state{{text-align:center;padding:60px 20px;color:{TEXT_SUB}}}
.empty-icon{{font-size:40px;color:#D1D5DB;margin-bottom:12px}}
.empty-title{{font-size:15px;font-weight:600;color:#374151;margin-bottom:6px}}
.empty-desc{{font-size:12px;color:{TEXT_SUB};line-height:1.8}}

.viz-control-bar{{background:{CARD_BG};border:1px solid {BORDER};border-radius:10px;padding:14px 18px;margin-bottom:16px;display:flex;align-items:flex-end;gap:20px;flex-wrap:wrap}}
.viz-control-item{{flex:1;min-width:180px}}
.viz-control-label{{font-size:11px;font-weight:700;color:{TEXT_SUB};text-transform:uppercase;letter-spacing:.07em;margin-bottom:5px}}
.viz-note{{font-size:11px;color:#9CA3AF;align-self:flex-end;padding-bottom:2px;display:flex;align-items:center;gap:4px}}

.dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner th{{background:{TBL_HDR}!important;color:white!important;font-size:12px!important;font-weight:600!important;text-align:center!important;padding:10px 8px!important;border:none!important;font-family:{FONT_BODY}!important;letter-spacing:.02em}}
.dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner td{{font-size:12.5px!important;text-align:center!important;padding:8px!important;border-color:#F3F4F6!important;font-family:{FONT_BODY}!important;color:{TEXT_MAIN}!important}}
.dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner tr:nth-child(even) td{{background:#FAFAFA!important}}
.dash-table-container .dash-spreadsheet-container .dash-spreadsheet-inner tr:hover td{{background:#EFF6FF!important;transition:background .1s}}

.Select-control{{border-radius:8px!important;border-color:{BORDER}!important;font-size:13px!important;font-family:{FONT_BODY}!important}}
.Select-control:hover{{border-color:#9CA3AF!important}}
.Select--is-focused .Select-control,.Select--is-focused .Select-control:hover{{border-color:{ACCENT}!important;box-shadow:0 0 0 3px rgba(59,130,246,.12)!important}}
.Select-menu-outer{{border-radius:8px!important;border-color:{BORDER}!important;box-shadow:0 8px 24px rgba(0,0,0,.1)!important;z-index:9999!important}}
.Select-option.is-selected{{background:{ACCENT}!important;color:white!important}}
.Select-option.is-focused{{background:#EFF6FF!important;color:{TEXT_MAIN}!important}}
.Select-multi-value-wrapper .Select-value{{background:#EFF6FF!important;border-color:#BFDBFE!important;border-radius:5px!important;color:{ACCENT_DARK}!important;font-size:12px!important}}
.Select-multi-value-wrapper .Select-value-icon{{border-color:#BFDBFE!important;color:{ACCENT}!important}}
.Select-multi-value-wrapper .Select-value-icon:hover{{background:#DBEAFE!important}}

/* 사이드바 드롭다운 오버라이드 */
.sb-select .Select-control{{background:#162030!important;border-color:#243650!important;color:#CBD5E1!important;min-height:34px!important;font-size:12px!important}}
.sb-select .Select-value-label,.sb-select .Select-single-value{{color:#CBD5E1!important}}
.sb-select .Select-placeholder{{color:#3D5470!important;font-size:12px}}
.sb-select .Select-input input{{color:#CBD5E1!important}}
.sb-select .Select-menu-outer{{background:#162030!important;border-color:#243650!important}}
.sb-select .Select-option{{color:#CBD5E1!important;font-size:12px;background:#162030!important}}
.sb-select .Select-option.is-focused{{background:#1E3050!important}}
.sb-select .Select-option.is-selected{{background:{ACCENT}!important;color:white!important}}
.sb-select .Select-arrow{{border-top-color:#64748B!important}}
.sb-select .Select-multi-value-wrapper .Select-value{{background:rgba(59,130,246,.2)!important;border-color:rgba(59,130,246,.35)!important;color:#93C5FD!important}}
.sb-select .Select-multi-value-wrapper .Select-value-icon{{border-color:rgba(59,130,246,.35)!important;color:#93C5FD!important}}
"""

# ─────────────────────────────────────────────
# 유틸 함수
# ─────────────────────────────────────────────
def parse_upload(contents, filename):
    _, content_string = contents.split(",")
    decoded = base64.b64decode(content_string)
    if filename.endswith(".csv"):
        return pd.read_csv(io.StringIO(decoded.decode("utf-8")))
    return pd.read_excel(io.BytesIO(decoded))

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

def sorted_levels(series):
    vals = list(series.dropna().astype(str).unique())
    try:
        return sorted(vals, key=lambda x: (0, float(x)))
    except Exception:
        return sorted(vals)

def fmt(val, mode, dec):
    return f"{val:.{dec}f}%" if mode == "pct" else f"{val:.{dec}f}"

def build_crosstab(df, t_var, c_vars, mode, dec):
    if t_var not in df.columns:
        return None
    t_levels = sorted_levels(df[t_var])
    fill = fmt(0, mode, dec)

    def get_vals(sub, level_col):
        rows = []
        for lvl in sorted_levels(sub[level_col]):
            grp = sub[sub[level_col].astype(str) == str(lvl)]
            total_w = grp["wt"].sum()
            row = {"구분": f"[{level_col}] {lvl}", "사례수(N)": f"{total_w:.{dec}f}"}
            for tl in t_levels:
                w = grp[grp[t_var].astype(str) == str(tl)]["wt"].sum()
                row[tl] = fmt(w / total_w * 100 if mode == "pct" and total_w > 0 else w, mode, dec)
            rows.append(row)
        return rows

    total_w = df["wt"].sum()
    overall = {"구분": "전체", "사례수(N)": f"{total_w:.{dec}f}"}
    for tl in t_levels:
        w = df[df[t_var].astype(str) == str(tl)]["wt"].sum()
        overall[tl] = fmt(w / total_w * 100 if mode == "pct" and total_w > 0 else w, mode, dec)

    all_rows = [overall]
    for c_var in c_vars:
        if c_var in df.columns:
            all_rows.extend(get_vals(df, c_var))

    return pd.DataFrame(all_rows, columns=["구분", "사례수(N)"] + t_levels).fillna(fill)

def make_table(result):
    return dash_table.DataTable(
        data=result.to_dict("records"),
        columns=[{"name": c, "id": c} for c in result.columns],
        page_size=200,
        style_table={"overflowX": "auto", "borderRadius": "6px", "overflow": "hidden"},
        style_cell={"fontFamily": FONT_BODY, "border": "1px solid #F3F4F6"},
        style_header={},
        style_data_conditional=[
            {"if": {"filter_query": '{구분} = "전체"'},
             "backgroundColor": "#EFF6FF", "fontWeight": "700", "color": TBL_HDR},
            {"if": {"column_id": "구분"}, "textAlign": "left", "paddingLeft": "16px"},
        ],
        filter_action="native",
        sort_action="native",
    )

# ─────────────────────────────────────────────
# 사이드바 헬퍼
# ─────────────────────────────────────────────
def sb_section(icon, title, children):
    return html.Div([
        html.Div([html.I(className=f"bi {icon}"), title], className="sb-section-title"),
        *children,
    ], className="sb-section")

def sb_label(text):
    return html.Div(text, className="sb-label")

def sb_dropdown(id_, opts, val, multi=False):
    return dcc.Dropdown(id=id_, options=opts, value=val, multi=multi,
                        clearable=False, className="sb-select",
                        style={"fontSize": "12px"})

# ─────────────────────────────────────────────
# 레이아웃
# ─────────────────────────────────────────────
sidebar = html.Div([
    html.Div([
        html.Div("Survey Analytics", className="logo-mark"),
        html.Div("결과 분석 대시보드", className="logo-title"),
        html.Div("Cross-tab & Visualization", className="logo-sub"),
    ], id="sidebar-logo"),

    sb_section("bi-cloud-upload", "데이터 업로드", [
        dcc.Upload(
            id="upload-file",
            children=html.Button(
                [html.I(className="bi bi-file-earmark-excel"), "  Excel 파일 선택"],
                className="upload-btn"),
            accept=".xlsx,.xls",
            max_size=100 * 1024 * 1024,
        ),
        html.Div(id="file-badge"),
    ]),

    sb_section("bi-pencil-square", "변수명 변경", [
        html.Div("기존변수=새변수  (줄바꿈으로 여러 개)", className="sb-label"),
        dbc.Textarea(id="var-names", placeholder="SQ1=지역\nSQ2=연령대",
                     rows=3, className="sb-textarea"),
    ]),

    sb_section("bi-tags", "값 라벨링", [
        html.Div("변수명: 1=값A, 2=값B", className="sb-label"),
        dbc.Textarea(id="labels",
                     placeholder="성별: 1=남자, 2=여자\n지역: 1=서울, 2=경기",
                     rows=4, className="sb-textarea"),
    ]),

    sb_section("bi-funnel", "가중치 & 필터", [
        sb_label("가중치 변수"),
        sb_dropdown("weight-var", [{"label": "(가중치 없음)", "value": "__none__"}], "__none__"),
        sb_label("필터 변수"),
        sb_dropdown("filter-var", [{"label": "(선택 안 함)", "value": "__none__"}], "__none__"),
        html.Div(id="filter-val-ui"),
        dcc.Dropdown(id="filter-vals", options=[], value=[], multi=True,
                     style={"display": "none"}, className="sb-select"),
    ]),

    sb_section("bi-sliders2", "출력 설정", [
        sb_label("표시 방식"),
        dbc.RadioItems(
            id="display-mode",
            options=[{"label": "비율 (%)", "value": "pct"},
                     {"label": "사례수 (N)", "value": "count"}],
            value="pct", inline=True, className="sb-radio",
        ),
        sb_label("소수점 자리수"),
        dbc.Input(id="decimals", type="number", value=1, min=0, max=5, step=1,
                  className="sb-number"),
    ]),
], id="sidebar")

main_content = html.Div([
    html.Div([
        html.Span([
            html.Span("분석 워크스페이스", className="page-title"),
            html.Span("교차분석표 & 시각화", className="page-sub"),
        ]),
    ], id="topbar"),

    html.Div([
        # 변수 설정 카드
        html.Div([
            html.Div([html.I(className="bi bi-grid-3x3-gap"), "분석 변수 설정"],
                     className="var-card-title"),
            dbc.Row([
                dbc.Col([
                    html.Div([html.Span("COL", className="badge-col"),
                              "가로(Column) 변수 — 주요 분석 문항"],
                             className="var-label"),
                    dcc.Dropdown(id="target-var", options=[],
                                 placeholder="분석 문항 선택 (다중)...", multi=True),
                ], md=6),
                dbc.Col([
                    html.Div([html.Span("ROW", className="badge-row"),
                              "세로(Row) 변수 — 집단 특성 문항"],
                             className="var-label"),
                    dcc.Dropdown(id="cross-vars", options=[],
                                 placeholder="집단 변수 선택 (다중)...", multi=True),
                ], md=6),
            ]),
        ], className="var-card"),

        # 탭
        dbc.Tabs([
            dbc.Tab(label="  데이터 테이블", tab_id="tab-table", children=[
                html.Div([
                    html.Button([html.I(className="bi bi-file-earmark-excel"), "  Excel"],
                                id="btn-excel", className="dl-btn dl-btn-excel"),
                    html.Button([html.I(className="bi bi-filetype-csv"), "  CSV"],
                                id="btn-csv", className="dl-btn dl-btn-csv"),
                    html.Span("문항별 시트/구분행으로 내보내기",
                              style={"fontSize": "11px", "color": "#9CA3AF", "marginLeft": "8px"}),
                ], style={"display": "flex", "alignItems": "center", "gap": "8px",
                          "marginBottom": "16px"}),
                dcc.Download(id="download-excel"),
                dcc.Download(id="download-csv"),
                html.Div(id="crosstab-table"),
            ]),
            dbc.Tab(label="  시각화", tab_id="tab-chart", children=[
                html.Div([
                    html.Div([
                        html.Div("시각화 문항", className="viz-control-label"),
                        dcc.Dropdown(id="plot-tvar", options=[],
                                     placeholder="문항 선택...", clearable=False),
                    ], className="viz-control-item"),
                    html.Div([
                        html.Div("그룹 기준", className="viz-control-label"),
                        dcc.Dropdown(id="plot-group",
                                     options=[{"label": "전체", "value": "전체"}],
                                     value="전체", clearable=False),
                    ], className="viz-control-item"),
                    html.Div([html.I(className="bi bi-info-circle"),
                              "1번 응답이 왼쪽/위부터 배치"],
                             className="viz-note"),
                ], className="viz-control-bar"),
                html.Div(
                    dcc.Graph(id="result-plot", style={"height": "520px"},
                              config={"displayModeBar": True, "displaylogo": False,
                                      "modeBarButtonsToRemove": ["select2d", "lasso2d"]}),
                    style={"background": CARD_BG, "borderRadius": "12px",
                           "border": f"1px solid {BORDER}", "padding": "8px",
                           "boxShadow": "0 1px 3px rgba(0,0,0,.04)"},
                ),
            ]),
        ], id="main-tabs", active_tab="tab-table"),
    ], id="content-area"),
], id="main-content")

app.index_string = f"""
<!DOCTYPE html>
<html>
<head>
{{%metas%}}
<title>{{%title%}}</title>
{{%favicon%}}
{{%css%}}
<style>
{GLOBAL_CSS}
</style>
</head>
<body>
{{%app_entry%}}
<footer>{{%config%}}{{%scripts%}}{{%renderer%}}</footer>
</body>
</html>
"""

app.layout = html.Div([
    dcc.Store(id="store-raw"),
    dcc.Store(id="store-renamed"),
    dcc.Store(id="store-labeled"),
    dcc.Store(id="store-processed"),
    dcc.Store(id="store-crosstab"),
    html.Div([sidebar, main_content],
             style={"display": "flex", "alignItems": "stretch", "minHeight": "100vh"}),
])


# ─────────────────────────────────────────────
# 콜백
# ─────────────────────────────────────────────
@app.callback(
    Output("store-raw", "data"),
    Output("file-badge", "children"),
    Input("upload-file", "contents"),
    State("upload-file", "filename"),
    prevent_initial_call=True,
)
def store_raw(contents, filename):
    if contents is None:
        return no_update, no_update
    df = parse_upload(contents, filename)
    badge = html.Div([html.I(className="bi bi-check-circle-fill",
                              style={"color": "#34D399"}), filename],
                     className="file-badge")
    return df.to_json(date_format="iso", orient="split"), badge


@app.callback(
    Output("store-renamed", "data"),
    Input("store-raw", "data"),
    Input("var-names", "value"),
    prevent_initial_call=True,
)
def store_renamed(raw_json, rename_text):
    if not raw_json:
        return no_update
    df = pd.read_json(io.StringIO(raw_json), orient="split")
    return apply_rename(df, rename_text).to_json(date_format="iso", orient="split")


@app.callback(
    Output("store-labeled", "data"),
    Input("store-renamed", "data"),
    Input("labels", "value"),
    prevent_initial_call=True,
)
def store_labeled(renamed_json, label_text):
    if not renamed_json:
        return no_update
    df = pd.read_json(io.StringIO(renamed_json), orient="split")
    return apply_labels(df, label_text).to_json(date_format="iso", orient="split")


@app.callback(
    Output("weight-var", "options"),
    Output("filter-var", "options"),
    Output("target-var", "options"),
    Output("cross-vars", "options"),
    Input("store-renamed", "data"),
    prevent_initial_call=True,
)
def update_dropdowns(renamed_json):
    if not renamed_json:
        return no_update, no_update, no_update, no_update
    df = pd.read_json(io.StringIO(renamed_json), orient="split")
    cols = df.columns.tolist()
    col_opts = [{"label": c, "value": c} for c in cols]
    return (
        [{"label": "(가중치 없음)", "value": "__none__"}] + col_opts,
        [{"label": "(선택 안 함)", "value": "__none__"}] + col_opts,
        col_opts,
        col_opts,
    )


@app.callback(
    Output("filter-val-ui", "children"),
    Output("filter-vals", "options"),
    Output("filter-vals", "value"),
    Output("filter-vals", "style"),
    Input("filter-var", "value"),
    Input("store-labeled", "data"),
    prevent_initial_call=True,
)
def filter_val_ui(f_var, labeled_json):
    hidden = {"display": "none"}
    if not labeled_json or not f_var or f_var == "__none__":
        return [], [], [], hidden
    df = pd.read_json(io.StringIO(labeled_json), orient="split")
    if f_var not in df.columns:
        return [], [], [], hidden
    choices = sorted_levels(df[f_var])
    opts = [{"label": c, "value": c} for c in choices]
    label = html.Div(f"'{f_var}' 포함 값", className="sb-label",
                     style={"marginTop": "8px"})
    return [label], opts, choices, {"display": "block"}


@app.callback(
    Output("store-processed", "data"),
    Input("store-labeled", "data"),
    Input("weight-var", "value"),
    Input("filter-var", "value"),
    Input("filter-vals", "value"),
    prevent_initial_call=True,
)
def store_processed(labeled_json, weight_var, filter_var, filter_vals):
    if not labeled_json:
        return no_update
    df = pd.read_json(io.StringIO(labeled_json), orient="split")
    if filter_var and filter_var != "__none__" and filter_vals and filter_var in df.columns:
        df = df[df[filter_var].astype(str).isin([str(v) for v in filter_vals])]
    if weight_var and weight_var != "__none__" and weight_var in df.columns:
        df["wt"] = pd.to_numeric(df[weight_var], errors="coerce").fillna(0)
    else:
        df["wt"] = 1.0
    return df.to_json(date_format="iso", orient="split")


@app.callback(
    Output("plot-group", "options"),
    Output("plot-group", "value"),
    Output("plot-tvar", "options"),
    Output("plot-tvar", "value"),
    Input("cross-vars", "value"),
    Input("target-var", "value"),
    prevent_initial_call=True,
)
def update_plot_controls(cross_vars, target_vars):
    group_opts = [{"label": "전체", "value": "전체"}]
    if cross_vars:
        group_opts += [{"label": v, "value": v} for v in cross_vars]
    tvar_opts, tvar_val = [], None
    if target_vars:
        tvar_opts = [{"label": v, "value": v} for v in target_vars]
        tvar_val = target_vars[0]
    return group_opts, "전체", tvar_opts, tvar_val


@app.callback(
    Output("crosstab-table", "children"),
    Output("store-crosstab", "data"),
    Input("store-processed", "data"),
    Input("target-var", "value"),
    Input("cross-vars", "value"),
    Input("display-mode", "value"),
    Input("decimals", "value"),
    prevent_initial_call=True,
)
def update_crosstab(proc_json, t_vars, c_vars, mode, dec):
    if not proc_json or not t_vars or not c_vars:
        return html.Div([
            html.Div(html.I(className="bi bi-bar-chart-line"), className="empty-icon"),
            html.Div("분석 변수를 선택하세요", className="empty-title"),
            html.Div("가로(COL) 변수와 세로(ROW) 변수를\n각각 하나 이상 선택하면\n교차 분석표가 생성됩니다.",
                     className="empty-desc", style={"whiteSpace": "pre-line"}),
        ], className="empty-state"), None

    dec = int(dec) if dec is not None else 1
    df = pd.read_json(io.StringIO(proc_json), orient="split")
    blocks, all_tables = [], {}

    for i, t_var in enumerate(t_vars, 1):
        result = build_crosstab(df, t_var, c_vars, mode, dec)
        if result is None:
            continue
        all_tables[t_var] = result
        blocks.append(html.Div([
            html.Div([
                html.Div(str(i), className="q-index"),
                html.Div(t_var, className="q-title"),
            ], className="result-card-header"),
            html.Div(make_table(result), className="result-card-body"),
        ], className="result-card"))

    if not blocks:
        return html.Div("분석 오류가 발생했습니다.", className="empty-state"), None

    store_payload = {k: v.to_json(orient="split") for k, v in all_tables.items()}
    return html.Div(blocks), str(store_payload)


@app.callback(
    Output("download-excel", "data"),
    Input("btn-excel", "n_clicks"),
    State("store-crosstab", "data"),
    prevent_initial_call=True,
)
def download_excel(n, crosstab_data):
    if not crosstab_data:
        return no_update
    try:
        payload = eval(crosstab_data)
    except Exception:
        return no_update
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({"bold": True, "bg_color": TBL_HDR, "font_color": "white",
                                  "align": "center", "border": 1, "font_size": 11})
        total_fmt = wb.add_format({"bold": True, "bg_color": "#EFF6FF", "font_color": TBL_HDR,
                                    "align": "center", "border": 1, "font_size": 11})
        for t_var, json_str in payload.items():
            df = pd.read_json(io.StringIO(json_str), orient="split")
            sheet_name = str(t_var)[:31]
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for col_num, col in enumerate(df.columns):
                ws.write(0, col_num, col, hdr_fmt)
            for row_num, row in enumerate(df.itertuples(index=False), 1):
                if str(row[0]) == "전체":
                    for col_num, val in enumerate(row):
                        ws.write(row_num, col_num, val, total_fmt)
            ws.set_column(0, 0, 30)
            ws.set_column(1, len(df.columns) - 1, 12)
            ws.freeze_panes(1, 0)
    buf.seek(0)
    return dcc.send_bytes(buf.read(), "교차분석표.xlsx")


@app.callback(
    Output("download-csv", "data"),
    Input("btn-csv", "n_clicks"),
    State("store-crosstab", "data"),
    prevent_initial_call=True,
)
def download_csv(n, crosstab_data):
    if not crosstab_data:
        return no_update
    try:
        payload = eval(crosstab_data)
    except Exception:
        return no_update
    parts = []
    for t_var, json_str in payload.items():
        df = pd.read_json(io.StringIO(json_str), orient="split")
        sep = pd.DataFrame([["▶ " + t_var] + [""] * (len(df.columns) - 1)], columns=df.columns)
        parts.extend([sep, df])
    combined = pd.concat(parts, ignore_index=True)
    return dcc.send_data_frame(combined.to_csv, "교차분석표.csv", index=False, encoding="utf-8-sig")


@app.callback(
    Output("result-plot", "figure"),
    Input("store-processed", "data"),
    Input("plot-tvar", "value"),
    Input("plot-group", "value"),
    Input("decimals", "value"),
    prevent_initial_call=True,
)
def update_plot(proc_json, t_var, p_group, dec):
    def empty_fig(msg="시각화할 문항을 선택하세요"):
        fig = go.Figure()
        fig.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            annotations=[dict(text=msg, x=0.5, y=0.5, xref="paper", yref="paper",
                              showarrow=False, font=dict(size=14, color="#9CA3AF",
                                                         family=FONT_BODY))],
            xaxis=dict(visible=False), yaxis=dict(visible=False),
            margin=dict(l=0, r=0, t=0, b=0),
        )
        return fig

    if not proc_json or not t_var:
        return empty_fig()

    dec = int(dec) if dec is not None else 1
    df = pd.read_json(io.StringIO(proc_json), orient="split")
    if t_var not in df.columns:
        return empty_fig()

    t_levels = sorted_levels(df[t_var])
    palette = ["#3B82F6","#10B981","#F59E0B","#EF4444","#8B5CF6",
               "#06B6D4","#F97316","#EC4899","#14B8A6","#6366F1"]

    base_layout = dict(
        paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family=FONT_BODY, size=12, color=TEXT_MAIN),
        margin=dict(l=60, r=30, t=55, b=60),
        hoverlabel=dict(bgcolor="white", bordercolor=BORDER,
                        font=dict(family=FONT_BODY, size=12)),
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
                marker=dict(color=palette[i % len(palette)], line=dict(width=0),
                            cornerradius=4),
                text=[f"<b>{row['pct']:.{dec}f}%</b>"],
                textposition="outside",
                hovertemplate=f"<b>{row[t_var]}</b><br>{row['pct']:.{dec}f}%<extra></extra>",
                showlegend=False, width=0.5,
            ))
        fig.update_layout(
            **base_layout,
            title=dict(text=f"<b>{t_var}</b>  전체 응답 비율",
                       font=dict(size=14, color=TBL_HDR, family=FONT_BODY),
                       x=0, xanchor="left", pad=dict(b=10)),
            xaxis=dict(title=None, showgrid=False, zeroline=False,
                       tickfont=dict(size=12)),
            yaxis=dict(title="비율 (%)", showgrid=True, gridcolor="#F3F4F6",
                       zeroline=False,
                       range=[0, (grp["pct"].max() if not grp.empty else 100) * 1.2]),
            bargap=0.4,
        )
        return fig

    # 교차 누적 가로 막대
    if p_group not in df.columns:
        return empty_fig()

    p_levels = sorted_levels(df[p_group])
    sub = df[df[t_var].notna() & df[p_group].notna()].copy()
    sub[t_var] = sub[t_var].astype(str)
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
            y=p_levels, x=vals, name=str(tl),
            orientation="h",
            marker=dict(color=palette[i % len(palette)], line=dict(width=0)),
            text=[f"<b>{v:.{dec}f}%</b>" if v > 4 else "" for v in vals],
            textposition="inside", insidetextanchor="middle",
            hovertemplate=f"<b>{tl}</b><br>%{{y}}: %{{x:.{dec}f}}%<extra></extra>",
        ))
    fig.update_layout(
        **base_layout,
        barmode="stack",
        title=dict(text=f"<b>{p_group}</b> 특성별  <b>{t_var}</b>  응답 분포",
                   font=dict(size=14, color=TBL_HDR, family=FONT_BODY),
                   x=0, xanchor="left", pad=dict(b=10)),
        xaxis=dict(title="비율 (%)", range=[0, 100], showgrid=True,
                   gridcolor="#F3F4F6", zeroline=False, ticksuffix="%"),
        yaxis=dict(title=None, showgrid=False, zeroline=False,
                   tickfont=dict(size=12)),
        legend=dict(title=dict(text=t_var, font=dict(size=12, color=TEXT_SUB)),
                    orientation="h", yanchor="bottom", y=-0.22,
                    xanchor="left", x=0,
                    font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
        bargap=0.25,
    )
    return fig


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True, port=8050)
