import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

st.set_page_config(
    page_title="EV Market Intelligence Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #0f1117; }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .dashboard-title {
        font-size: 2.4rem; font-weight: 800;
        text-align: center; margin-bottom: 0.2rem;
        background: linear-gradient(90deg, #00d4ff, #00ff88);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .dashboard-subtitle {
        text-align: center; color: #8892a4; font-size: 1rem; margin-bottom: 1.5rem;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1a1f2e, #252b3b);
        border: 1px solid #2d3548; border-radius: 16px;
        padding: 1.4rem 1.6rem; text-align: center;
    }
    .kpi-label  { color: #8892a4; font-size: 0.82rem; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 0.4rem; }
    .kpi-value  { color: #ffffff; font-size: 2rem; font-weight: 800; }
    .kpi-delta  { font-size: 0.8rem; margin-top: 0.3rem; }
    .kpi-delta.positive { color: #00ff88; }
    .kpi-delta.neutral  { color: #00d4ff; }
    .section-header {
        color: #ffffff; font-size: 1.2rem; font-weight: 700;
        border-left: 4px solid #00d4ff; padding-left: 0.8rem;
        margin: 1.5rem 0 0.8rem 0;
    }
    .section-subtext { color: #8892a4; font-size: 0.85rem; margin-bottom: 1rem; }
    /* Key Insights banner */
    .insights-banner {
        background: linear-gradient(135deg, #0d1b30, #0a2218);
        border: 1px solid #1e3a55;
        border-radius: 16px;
        padding: 1.2rem 1.6rem;
        margin-bottom: 1.4rem;
    }
    .insights-banner-title {
        color: #00d4ff; font-size: 1rem; font-weight: 700;
        text-transform: uppercase; letter-spacing: 1.5px;
        margin-bottom: 0.8rem;
    }
    .insights-grid {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 1rem;
    }
    .insight-item {
        background: rgba(255,255,255,0.03);
        border-left: 3px solid #00d4ff;
        border-radius: 0 8px 8px 0;
        padding: 0.7rem 0.9rem;
        font-size: 1rem; line-height: 1.6; color: #a8cce0;
    }
    .insight-item b { color: #00ff88; }
    .insight-item .label {
        font-size: 1rem; font-weight: 700; color: #00d4ff;
        text-transform: uppercase; letter-spacing: 1px;
        margin-bottom: 0.3rem;
    }
    /* Recommended Action cards */
    .action-card {
        background: linear-gradient(135deg, #141e2e, #1a2a1a);
        border: 1px solid #2a4a2a;
        border-radius: 14px;
        padding: 1.2rem 1.3rem;
    }
    .action-priority {
        display: inline-block;
        font-size: 0.65rem; font-weight: 800; letter-spacing: 1.5px;
        text-transform: uppercase; padding: 0.2rem 0.6rem;
        border-radius: 20px; margin-bottom: 0.6rem;
    }
    .priority-immediate { background: rgba(255,68,102,0.2); color: #ff6688; border: 1px solid #ff4466; }
    .priority-nearterm  { background: rgba(255,170,68,0.2); color: #ffcc66; border: 1px solid #ffaa44; }
    .priority-strategic { background: rgba(0,212,255,0.2);  color: #66ddff; border: 1px solid #00d4ff; }
    .action-title { font-size: 1.05rem; font-weight: 700; color: #ffffff; margin-bottom: 0.5rem; }
    .action-body  { font-size: 1rem; color: #8aacbe; line-height: 1.65; }
    .action-body b { color: #00ff88; }
    .action-metric {
        margin-top: 0.7rem; padding-top: 0.6rem;
        border-top: 1px solid #2a3a2a;
        font-size: 0.85rem; color: #5a7a5a;
    }
    .action-metric b { color: #00ff88; }
</style>
""", unsafe_allow_html=True)


# ── Data Loading ─────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    base = os.path.dirname(os.path.abspath(__file__))

    def find_file(pattern):
        for f in os.listdir(base):
            if pattern in f and f.endswith(".xlsx"):
                return os.path.join(base, f)
        raise FileNotFoundError(
            f"Could not find a file matching '*{pattern}*.xlsx' in:\n{base}\n\n"
            "Make sure all three .xlsx files are in the same folder as this script."
        )

    trips_path = find_file("10318")
    ev_path    = find_file("10962")
    ch_path    = find_file("10972")

    # Trips by length
    trips_raw  = pd.read_excel(trips_path, sheet_name="Condensed", header=None)
    header_row = next((i for i, row in trips_raw.iterrows() if str(row.iloc[1]).strip() == "Miles"), None)
    if header_row is None:
        raise ValueError("Could not find header row in trips file.")
    trips = trips_raw.iloc[header_row + 1:, [1, 2]].copy()
    trips.columns = ["Miles", "Share"]
    trips = trips.dropna(subset=["Miles"])
    trips["Miles"] = trips["Miles"].astype(str).str.strip()
    trips = trips[trips["Miles"] != "Miles"]
    trips["Share"] = pd.to_numeric(trips["Share"], errors="coerce")
    trips = trips.dropna(subset=["Share"])
    trips["Share_pct"] = (trips["Share"] * 100).round(1)
    trips = trips.reset_index(drop=True)

    # EV registrations by state
    ev_raw     = pd.read_excel(ev_path, sheet_name="Condensed", header=None)
    header_row = next((i for i, row in ev_raw.iterrows() if str(row.iloc[1]).strip() == "State"), None)
    if header_row is None:
        raise ValueError("Could not find header row in EV registrations file.")
    ev = ev_raw.iloc[header_row + 1:, [1, 2]].copy()
    ev.columns = ["State", "Registrations"]
    ev = ev.dropna(subset=["State"])
    ev["State"] = ev["State"].astype(str).str.strip()
    ev = ev[ev["State"] != "State"]
    ev["Registrations"] = pd.to_numeric(ev["Registrations"], errors="coerce")
    ev = ev.dropna(subset=["Registrations"])
    ev = ev.reset_index(drop=True)

    # Charging infrastructure
    ch_raw     = pd.read_excel(ch_path, sheet_name="Condensed", header=None)
    header_row = next((i for i, row in ch_raw.iterrows() if str(row.iloc[1]).strip() == "Year"), None)
    if header_row is None:
        raise ValueError("Could not find header row in charging infrastructure file.")
    ch = ch_raw.iloc[header_row + 1:, [1, 2, 3]].copy()
    ch.columns = ["Year", "Ports", "Stations"]
    ch = ch.dropna(subset=["Year"])
    ch["Year"]     = pd.to_numeric(ch["Year"],     errors="coerce")
    ch["Ports"]    = pd.to_numeric(ch["Ports"],    errors="coerce")
    ch["Stations"] = pd.to_numeric(ch["Stations"], errors="coerce")
    ch = ch.dropna()
    ch["Year"] = ch["Year"].astype(int)
    ch = ch.reset_index(drop=True)

    return trips, ev, ch


try:
    trips, ev, ch = load_data()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"Unexpected error while loading data: {e}")
    st.stop()

# ── Derived KPIs ─────────────────────────────────────────────────────────────
total_ev       = int(ev["Registrations"].sum())
total_stations = int(ch["Stations"].iloc[-1])
total_ports    = int(ch["Ports"].iloc[-1])
short_trip_pct = trips[trips["Miles"].isin(["<6", "6–10"])]["Share_pct"].sum()
yoy_ports      = ((ch["Ports"].iloc[-1] / ch["Ports"].iloc[-2]) - 1) * 100

STATE_ABBR = {
    "Alabama":"AL","Alaska":"AK","Arizona":"AZ","Arkansas":"AR","California":"CA",
    "Colorado":"CO","Connecticut":"CT","Delaware":"DE","District of Columbia":"DC",
    "Florida":"FL","Georgia":"GA","Hawaii":"HI","Idaho":"ID","Illinois":"IL",
    "Indiana":"IN","Iowa":"IA","Kansas":"KS","Kentucky":"KY","Louisiana":"LA",
    "Maine":"ME","Maryland":"MD","Massachusetts":"MA","Michigan":"MI","Minnesota":"MN",
    "Mississippi":"MS","Missouri":"MO","Montana":"MT","Nebraska":"NE","Nevada":"NV",
    "New Hampshire":"NH","New Jersey":"NJ","New Mexico":"NM","New York":"NY",
    "North Carolina":"NC","North Dakota":"ND","Ohio":"OH","Oklahoma":"OK","Oregon":"OR",
    "Pennsylvania":"PA","Rhode Island":"RI","South Carolina":"SC","South Dakota":"SD",
    "Tennessee":"TN","Texas":"TX","Utah":"UT","Vermont":"VT","Virginia":"VA",
    "Washington":"WA","West Virginia":"WV","Wisconsin":"WI","Wyoming":"WY"
}
PLOT_BG    = "rgba(0,0,0,0)"
AXIS_COLOR = "#8892a4"
GRID_COLOR = "#1e2535"

# ════════════════════════════════════════════════════════════════════════════
# HEADER
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="dashboard-title">⚡ EV Market Intelligence Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="dashboard-subtitle">U.S. EV Market Overview · Charging Infrastructure · Driving Behavior Insights</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# CHANGE 1 — KEY INSIGHTS BANNER
# ════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="insights-banner">
  <div class="insights-banner-title">💡 Key Insights</div>
  <div class="insights-grid">
    <div class="insight-item">
      <div class="label">Infrastructure</div>
      U.S. public charging ports surged from <b>417 (2007)</b> to <b>168,388 (2023)</b> —
      a <b>400× increase</b> in 16 years, actively eliminating range anxiety nationwide.
    </div>
    <div class="insight-item">
      <div class="label">Market Landscape</div>
      California alone holds <b>~40%</b> of all U.S. EV registrations.
      The top 10 states capture <b>75%+</b> of the market, leaving enormous
      untapped growth potential across the remaining 40+ states.
    </div>
    <div class="insight-item">
      <div class="label">Driving Behavior</div>
      <b>71.4%</b> of all U.S. daily trips are 10 miles or less
      (Oak Ridge NHTS 2022) — well within every modern EV's range,
      confirming that EVs already meet the majority of American driving needs.
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════════════════════
k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Total EV Registrations (2023)</div>
        <div class="kpi-value">{total_ev/1_000_000:.2f}M</div>
        <div class="kpi-delta positive">↑ Across all 51 states</div>
    </div>""", unsafe_allow_html=True)
with k2:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Public Charging Stations</div>
        <div class="kpi-value">{total_stations:,}</div>
        <div class="kpi-delta positive">↑ Latest 2023 data</div>
    </div>""", unsafe_allow_html=True)
with k3:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Total Charging Ports</div>
        <div class="kpi-value">{total_ports:,}</div>
        <div class="kpi-delta positive">↑ YoY +{yoy_ports:.1f}%</div>
    </div>""", unsafe_allow_html=True)
with k4:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">Short Trips Share (≤10 mi)</div>
        <div class="kpi-value">{short_trip_pct:.0f}%</div>
        <div class="kpi-delta neutral">→ Perfectly suited for EVs</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# ROW 1 — Choropleth Map + Top 10 Bar
# ════════════════════════════════════════════════════════════════════════════
col_left, col_right = st.columns([3, 2])

with col_left:
    st.markdown('<div class="section-header">🗺️ EV Registration Heatmap by State</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtext">Darker shading = higher EV adoption. Hover for state details.</div>', unsafe_allow_html=True)
    ev_map = ev.copy()
    ev_map["Code"] = ev_map["State"].map(STATE_ABBR)
    ev_map = ev_map.dropna(subset=["Code"])
    fig_map = px.choropleth(
        ev_map, locations="Code", locationmode="USA-states",
        color="Registrations", scope="usa",
        color_continuous_scale=[[0,"#0d1b2a"],[0.2,"#1a4a6e"],[0.5,"#0088cc"],[0.8,"#00ccff"],[1,"#00ff88"]],
        hover_name="State",
        hover_data={"Registrations": ":,", "Code": False},
        labels={"Registrations": "EV Registrations"},
    )
    fig_map.update_layout(
        paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG,
        geo=dict(bgcolor=PLOT_BG, lakecolor=PLOT_BG, landcolor="#1a1f2e", subunitcolor="#2d3548"),
        coloraxis_colorbar=dict(title=dict(text="Registrations", font=dict(color=AXIS_COLOR)), tickfont=dict(color=AXIS_COLOR)),
        margin=dict(l=0, r=0, t=0, b=0), height=340,
    )
    st.plotly_chart(fig_map, use_container_width=True)

with col_right:
    st.markdown('<div class="section-header">🏆 Top 10 States by EV Registrations</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtext">California leads with ~40% of all U.S. EV registrations</div>', unsafe_allow_html=True)
    top10 = ev.sort_values("Registrations", ascending=True).tail(10)
    fig_bar = go.Figure(go.Bar(
        x=top10["Registrations"], y=top10["State"], orientation="h",
        marker=dict(color=top10["Registrations"], colorscale=[[0,"#1a4a6e"],[1,"#00ff88"]], showscale=False),
        text=top10["Registrations"].apply(lambda x: f"{x/1_000_000:.2f}M" if x >= 1_000_000 else f"{x/1000:.0f}K"),
        textposition="outside", textfont=dict(color="#ffffff", size=11),
        hovertemplate="%{y}: %{x:,} registrations<extra></extra>",
    ))
    fig_bar.update_layout(
        paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG,
        xaxis=dict(showgrid=False, showticklabels=False, color=AXIS_COLOR),
        yaxis=dict(color=AXIS_COLOR, tickfont=dict(size=11), gridcolor=GRID_COLOR),
        margin=dict(l=10, r=70, t=10, b=10), height=340,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# ROW 2 — Charging Infrastructure Growth
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">📈 U.S. Public EV Charging Infrastructure Growth (2007–2023)</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtext">Charging ports grew from 417 to 168,388 — a 400× increase over 16 years</div>', unsafe_allow_html=True)

fig_ch = make_subplots(specs=[[{"secondary_y": True}]])
fig_ch.add_trace(go.Scatter(
    x=ch["Year"], y=ch["Ports"], name="Charging Ports",
    mode="lines+markers", line=dict(color="#00d4ff", width=3),
    marker=dict(size=6, color="#00d4ff"),
    fill="tozeroy", fillcolor="rgba(0,212,255,0.08)",
), secondary_y=False)
fig_ch.add_trace(go.Bar(
    x=ch["Year"], y=ch["Stations"], name="Charging Stations",
    marker_color="rgba(0,255,136,0.5)", marker_line_color="#00ff88", marker_line_width=1,
), secondary_y=True)
fig_ch.add_annotation(
    x=ch["Year"].iloc[-1], y=ch["Ports"].iloc[-1],
    text=f"<b>{ch['Ports'].iloc[-1]:,}</b> ports",
    showarrow=True, arrowhead=2, arrowcolor="#00d4ff",
    font=dict(color="#00d4ff", size=12), ax=-80, ay=-40,
)
fig_ch.update_layout(
    paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG,
    legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(color=AXIS_COLOR)),
    xaxis=dict(color=AXIS_COLOR, gridcolor=GRID_COLOR, tickmode="linear", dtick=2),
    yaxis=dict(color=AXIS_COLOR, gridcolor=GRID_COLOR, title="Charging Ports"),
    yaxis2=dict(color=AXIS_COLOR, gridcolor="rgba(0,0,0,0)", title="Charging Stations"),
    margin=dict(l=10, r=10, t=40, b=10), height=300, hovermode="x unified",
)
st.plotly_chart(fig_ch, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# ROW 3 — Trip Length + Donut
# ════════════════════════════════════════════════════════════════════════════
col_a, col_b = st.columns([3, 2])

with col_a:
    st.markdown('<div class="section-header">🚗 Distribution of U.S. Daily Trip Lengths</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtext">Most trips are well within EV range — EVs handle everyday commuting with ease</div>', unsafe_allow_html=True)
    is_short = trips["Miles"].isin(["<6", "6–10"])
    fig_trips = go.Figure(go.Bar(
        x=trips["Miles"], y=trips["Share_pct"],
        marker_color=["#00ff88" if s else "#2d3548" for s in is_short],
        marker_line_color=["#00ff88" if s else "#3d4560" for s in is_short],
        marker_line_width=1.5,
        text=trips["Share_pct"].apply(lambda x: f"{x:.1f}%"),
        textposition="outside", textfont=dict(color="#ffffff", size=12),
        hovertemplate="%{x} miles: %{y}%<extra></extra>",
    ))
    fig_trips.add_shape(type="line", x0=-0.4, x1=1.4, y0=60, y1=60,
                        line=dict(color="#00ff88", width=1, dash="dot"))
    fig_trips.add_annotation(x=0.5, y=63, text="<b>71.4% short trips</b> (ideal EV use case)",
                              showarrow=False, font=dict(color="#00ff88", size=12))
    fig_trips.update_layout(
        paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG,
        xaxis=dict(color=AXIS_COLOR, title="Trip Length (Miles)", gridcolor=GRID_COLOR),
        yaxis=dict(color=AXIS_COLOR, title="Share of All Trips (%)", gridcolor=GRID_COLOR),
        margin=dict(l=10, r=10, t=40, b=10), height=320, showlegend=False,
    )
    st.plotly_chart(fig_trips, use_container_width=True)

with col_b:
    st.markdown('<div class="section-header">🍩 Short vs. Medium vs. Long Trips</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtext">Over 70% of trips are under 10 miles</div>', unsafe_allow_html=True)
    short = trips[trips["Miles"].isin(["<6", "6–10"])]["Share_pct"].sum()
    mid   = trips[trips["Miles"].isin(["11–15", "16–20"])]["Share_pct"].sum()
    long_ = trips[trips["Miles"].isin(["21–30", ">30"])]["Share_pct"].sum()
    fig_donut = go.Figure(go.Pie(
        labels=["Short ≤10 mi", "Medium 11–20 mi", "Long >20 mi"],
        values=[round(short, 1), round(mid, 1), round(long_, 1)],
        hole=0.62,
        marker=dict(colors=["#00ff88","#00d4ff","#2d3548"], line=dict(color="#0f1117", width=3)),
        textinfo="label+percent", textfont=dict(size=12, color="#ffffff"),
        hoverinfo="label+value+percent",
    ))
    fig_donut.add_annotation(
        text=f"<b>{short:.0f}%</b><br>Short",
        x=0.5, y=0.5, showarrow=False, font=dict(size=18, color="#00ff88"),
    )
    fig_donut.update_layout(
        paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG, showlegend=False,
        margin=dict(l=10, r=10, t=10, b=10), height=320,
    )
    st.plotly_chart(fig_donut, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# YoY chart now full-width
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">📊 Charging Port Year-over-Year Growth Rate</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtext">After explosive early growth, expansion has entered a steady, maturing phase — signaling reliable long-term infrastructure momentum</div>', unsafe_allow_html=True)

ch_yoy = ch.copy()
ch_yoy["YoY_Ports"] = ch_yoy["Ports"].pct_change() * 100
ch_yoy = ch_yoy.dropna()

fig_yoy = go.Figure(go.Bar(
    x=ch_yoy["Year"], y=ch_yoy["YoY_Ports"].round(1),
    marker_color=["#00ff88" if v >= 0 else "#ff4466" for v in ch_yoy["YoY_Ports"]],
    text=ch_yoy["YoY_Ports"].round(0).astype(int).astype(str) + "%",
    textposition="outside", textfont=dict(color="#ffffff", size=10),
    hovertemplate="%{x} YoY: %{y:.1f}%<extra></extra>",
))
fig_yoy.update_layout(
    paper_bgcolor=PLOT_BG, plot_bgcolor=PLOT_BG,
    xaxis=dict(color=AXIS_COLOR, tickmode="linear", dtick=1, gridcolor=GRID_COLOR),
    yaxis=dict(color=AXIS_COLOR, title="YoY Growth Rate (%)", gridcolor=GRID_COLOR),
    margin=dict(l=10, r=10, t=30, b=10), height=280, showlegend=False,
)
st.plotly_chart(fig_yoy, use_container_width=True)

# ════════════════════════════════════════════════════════════════════════════
# EV Suitability Calculator
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">⚡ Interactive EV Suitability Calculator</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtext">Adjust the inputs below to see how well an EV fits your lifestyle</div>', unsafe_allow_html=True)

sim_col1, sim_col2, sim_col3 = st.columns(3)
with sim_col1:
    daily_miles = st.slider("Daily Driving Distance (miles)", 0, 200, 25)
with sim_col2:
    charge_access = st.selectbox("Charging Access", [
        "Can charge at home",
        "Public charger nearby",
        "Need to travel far to charge",
    ])
with sim_col3:
    drive_env = st.selectbox("Primary Driving Environment", [
        "Urban city commuting",           # stop-and-go → regen braking → EV most efficient
        "Mixed city & suburban",          # moderate trips → neutral baseline
        "Mostly highway / long distance", # sustained high speed → reduces real-world range
    ])

# Scoring: three distinct outcomes per driving environment
score = 100

if daily_miles > 120:  score -= 45
elif daily_miles > 80: score -= 30
elif daily_miles > 50: score -= 15

if charge_access == "Need to travel far to charge": score -= 25
elif charge_access == "Public charger nearby":      score -= 8
# "Can charge at home" → no penalty

if drive_env == "Urban city commuting":
    score += 5    # EVs are most efficient here: regen braking, lower average speed
elif drive_env == "Mixed city & suburban":
    score += 0    # neutral
elif drive_env == "Mostly highway / long distance":
    score -= 15   # sustained 65–75 mph cuts real-world range 15–25% vs EPA rating

score = max(0, min(100, score))
gauge_color = "#00ff88" if score >= 70 else "#ffaa44" if score >= 40 else "#ff4466"

fig_gauge = go.Figure(go.Indicator(
    mode="gauge+number", value=score,
    title={"text": "EV Suitability Score", "font": {"color": "#ffffff", "size": 16}},
    number={"font": {"color": gauge_color, "size": 40}, "suffix": "/100"},
    gauge={
        "axis": {"range": [0, 100], "tickcolor": AXIS_COLOR, "tickfont": {"color": AXIS_COLOR}},
        "bar": {"color": gauge_color, "thickness": 0.25},
        "bgcolor": "#1a1f2e", "bordercolor": "#2d3548",
        "steps": [
            {"range": [0,  40],  "color": "rgba(255,68,102,0.15)"},
            {"range": [40, 70],  "color": "rgba(255,170,68,0.15)"},
            {"range": [70, 100], "color": "rgba(0,255,136,0.15)"},
        ],
        "threshold": {"line": {"color": gauge_color, "width": 3}, "thickness": 0.75, "value": score},
    }
))
fig_gauge.update_layout(paper_bgcolor=PLOT_BG, height=230, margin=dict(l=30, r=30, t=30, b=10))

env_note = {
    "Urban city commuting":
        "City driving is where EVs truly shine — regenerative braking recovers energy at every stop, boosting real-world efficiency beyond the EPA rating.",
    "Mixed city & suburban":
        "Mixed driving is well within EV capabilities; your real-world range will closely match the manufacturer's estimate.",
    "Mostly highway / long distance":
        "Note: sustained highway speeds (65–75 mph) typically reduce real-world EV range by 15–25% vs. the EPA rating — factor this in when selecting a model.",
}[drive_env]

g1, g2 = st.columns([1, 2])
with g1:
    st.plotly_chart(fig_gauge, use_container_width=True)
with g2:
    if score >= 70:
        st.success(
            f"🟢 **Great fit for an EV!** Your daily {daily_miles}-mile drive is comfortably within the range of "
            f"most EVs (300–500 miles). "
            + ("Home charging is the most cost-effective option — charge overnight, drive all day. " if charge_access == "Can charge at home"
               else "Nearby public chargers are more than enough for your routine. ")
            + env_note + " Consider the **Tesla Model 3**, **Hyundai Ioniq 6**, or **Chevy Equinox EV**."
        )
    elif score >= 40:
        st.warning(
            f"🟡 **Manageable with some planning.** Your {daily_miles} miles/day is on the higher end but still workable. "
            f"We recommend a model with **400+ miles of EPA range** (e.g., Tesla Model Y Long Range) "
            f"and planning charging stops on longer trips. " + env_note
            + " A **plug-in hybrid (PHEV)** is also a solid transitional option."
        )
    else:
        st.error(
            f"🔴 **EVs may be challenging for your current situation.** Your daily {daily_miles} miles "
            f"combined with limited charging access makes a pure EV less practical right now. " + env_note
            + " Consider: ① a **PHEV** as a bridge, ② waiting for more local chargers, or "
            "③ revisiting as next-gen solid-state battery EVs arrive with greater range."
        )

# ════════════════════════════════════════════════════════════════════════════
# RECOMMENDED ACTIONS FOR EXECUTIVES
# ════════════════════════════════════════════════════════════════════════════
st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="section-header">🎯 Recommended Actions</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtext">A 20% increase in EV development & production spending, deployed across three strategic priorities</div>', unsafe_allow_html=True)

ra1, ra2, ra3 = st.columns(3)

with ra1:
    st.markdown("""
    <div class="action-card">
        <div class="action-priority priority-immediate">🔴 Priority 1 — R&D & Battery Architecture</div>
        <div class="action-title">Right-Size Batteries for Real-World Driving</div>
        <div class="action-body">
            <b>Key action:</b> Redirect R&D investment toward battery architectures
            optimized for the <b>sub-10-mile daily trip majority</b> rather than
            over-engineering for extreme range scenarios. Reducing per-unit battery
            capacity lowers production costs and accelerates volume scaling.
        </div>
        <div class="action-metric">📌 Data: <b>71.4%</b> of U.S. trips are ≤10 mi — today's EVs are already overbuilt for most drivers</div>
    </div>
    """, unsafe_allow_html=True)

with ra2:
    st.markdown("""
    <div class="action-card">
        <div class="action-priority priority-immediate">🔴 Priority 2 — Production Scaling</div>
        <div class="action-title">Ramp Output to Match Infrastructure Capacity</div>
        <div class="action-body">
            <b>Key action:</b> Expand physical production capacity to align with
            the <b>168,000+ public charging ports</b> already deployed nationwide.
            The charging network is ready — failing to scale output now means
            leaving subsidized demand on the table while competitors move in.
        </div>
        <div class="action-metric">📌 Context: Charging ports grew <b>+23.5% YoY</b> in 2023 — infrastructure is outpacing vehicle supply</div>
    </div>
    """, unsafe_allow_html=True)

with ra3:
    st.markdown("""
    <div class="action-card">
        <div class="action-priority priority-nearterm">🟡 Priority 3 — Regional Market Focus</div>
        <div class="action-title">Go Deep in CA, FL, TX, NY & WA First</div>
        <div class="action-body">
            <b>Key action:</b> Concentrate production deployment and marketing spend
            in the five highest-adoption states — <b>California, Florida, Texas,
            New York, and Washington</b> — where EV demand is already proven and
            infrastructure is established. Precision targeting maximizes immediate ROI
            over a spread-thin national approach.
        </div>
        <div class="action-metric">📌 Opportunity: These 5 states represent the largest validated EV demand hubs in the U.S.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(
    "<br><div style='text-align:center; color:#3d4560; font-size:0.78rem;'>"
    "Sources: Oak Ridge National Laboratory · U.S. DOE Alternative Fuels Station Locator · State DMV Registration Data (2023)"
    "</div>",
    unsafe_allow_html=True
)
