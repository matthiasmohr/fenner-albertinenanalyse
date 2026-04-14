import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="Labor-Analyse Albertinen-KH", layout="wide")

# --- Password protection ---
def check_password():
    correct = os.environ.get("APP_PASSWORD", "")
    if not correct:
        st.error("APP_PASSWORD ist nicht gesetzt.")
        st.stop()

    if st.session_state.get("authenticated"):
        return True

    st.title("Labor-Anforderungsanalyse — Albertinen-Krankenhaus")
    st.subheader("Anmeldung")
    pw = st.text_input("Passwort", type="password", key="pw_input")
    if st.button("Anmelden"):
        if pw == correct:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Falsches Passwort.")
    return False

if not check_password():
    st.stop()

st.title("Labor-Anforderungsanalyse — Albertinen-Krankenhaus")

# --- Data Loading ---
@st.cache_data
def load_data():
    df = pd.read_excel("input/input.xlsx")
    df.columns = ["Einsender", "Anforderung", "Untersuchungen", "GOÄ-Punkte"]
    return df

df = load_data()

# --- Global filters ---
metric = st.sidebar.radio("Kennzahl", ["Untersuchungen", "GOÄ-Punkte"], index=0)
exclude_zero_goae = st.sidebar.checkbox(
    "Ohne GOÄ-Punkte ausschließen",
    value=True,
    help="Einträge mit 0 GOÄ-Punkten ausblenden (i.d.R. interne Steuerkennzeichen und Statistiken)",
)
if exclude_zero_goae:
    df = df[df["GOÄ-Punkte"] > 0]
st.sidebar.markdown("---")
st.sidebar.markdown(f"**{df['Einsender'].nunique()}** Einsender · **{df['Anforderung'].nunique()}** Anforderungen")
st.sidebar.markdown(f"**{df['Untersuchungen'].sum():,.0f}** Untersuchungen · **{df['GOÄ-Punkte'].sum():,.0f}** GOÄ-Punkte")

# ============================================================
# 1. ÜBERBLICK
# ============================================================
st.header("1 · Überblick")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Einsender", df["Einsender"].nunique())
col2.metric("Anforderungen", df["Anforderung"].nunique())
col3.metric("Σ Untersuchungen", f"{df['Untersuchungen'].sum():,.0f}")
col4.metric("Σ GOÄ-Punkte", f"{df['GOÄ-Punkte'].sum():,.0f}")

einsender_agg = df.groupby("Einsender")[metric].sum().sort_values(ascending=False).reset_index()

fig_bar = px.bar(
    einsender_agg, x="Einsender", y=metric,
    title=f"{metric} je Einsender",
    color=metric, color_continuous_scale="Blues",
)
fig_bar.update_layout(xaxis_tickangle=-45, height=500, showlegend=False)
st.plotly_chart(fig_bar, use_container_width=True)

# ============================================================
# 2. TOP 30 ANFORDERUNGEN JE EINSENDER
# ============================================================
st.header("2 · Top 30 Anforderungen je Einsender")

einsender_list = einsender_agg["Einsender"].tolist()
selected_einsender = st.selectbox("Einsender auswählen", einsender_list, key="top30_einsender")

df_filtered = df[df["Einsender"] == selected_einsender]
top30_anf = df_filtered.nlargest(30, metric)

col_a, col_b = st.columns([2, 1])
with col_a:
    fig_top30_anf = px.bar(
        top30_anf, x=metric, y="Anforderung", orientation="h",
        title=f"Top 30 Anforderungen — {selected_einsender}",
        color=metric, color_continuous_scale="Teal",
    )
    fig_top30_anf.update_layout(yaxis=dict(autorange="reversed"), height=max(500, len(top30_anf) * 22), showlegend=False)
    st.plotly_chart(fig_top30_anf, use_container_width=True)
with col_b:
    st.dataframe(
        top30_anf[["Anforderung", "Untersuchungen", "GOÄ-Punkte"]].reset_index(drop=True),
        use_container_width=True, height=max(500, len(top30_anf) * 22),
    )

# ============================================================
# 3. TOP 30 EINSENDER JE ANFORDERUNG
# ============================================================
st.header("3 · Top 30 Einsender je Anforderung")

anforderung_agg = df.groupby("Anforderung")[metric].sum().sort_values(ascending=False)
anforderung_list = anforderung_agg.index.tolist()
selected_anforderung = st.selectbox("Anforderung auswählen", anforderung_list, key="top30_anforderung")

df_anf = df[df["Anforderung"] == selected_anforderung].nlargest(30, metric)

col_c, col_d = st.columns([2, 1])
with col_c:
    fig_top30_ein = px.bar(
        df_anf, x=metric, y="Einsender", orientation="h",
        title=f"Top 30 Einsender — {selected_anforderung}",
        color=metric, color_continuous_scale="Oranges",
    )
    fig_top30_ein.update_layout(yaxis=dict(autorange="reversed"), height=max(500, len(df_anf) * 22), showlegend=False)
    st.plotly_chart(fig_top30_ein, use_container_width=True)
with col_d:
    st.dataframe(
        df_anf[["Einsender", "Untersuchungen", "GOÄ-Punkte"]].reset_index(drop=True),
        use_container_width=True, height=max(500, len(df_anf) * 22),
    )

# ============================================================
# 4. MEKKO-CHART
# ============================================================
st.header("4 · Mekko-Chart (Einsender × Anforderungsgruppen)")

n_top_einsender = st.slider("Anzahl Top-Einsender", 5, 20, 10, key="mekko_n")
n_top_anf = st.slider("Anzahl Top-Anforderungsgruppen", 5, 15, 8, key="mekko_anf")

top_einsender = df.groupby("Einsender")[metric].sum().nlargest(n_top_einsender).index.tolist()
top_anforderungen = df.groupby("Anforderung")[metric].sum().nlargest(n_top_anf).index.tolist()

df_mekko = df.copy()
df_mekko["Anf_group"] = df_mekko["Anforderung"].where(df_mekko["Anforderung"].isin(top_anforderungen), "Sonstige")
df_mekko["Ein_group"] = df_mekko["Einsender"].where(df_mekko["Einsender"].isin(top_einsender), "Sonstige")

pivot = df_mekko.groupby(["Ein_group", "Anf_group"])[metric].sum().reset_index()
einsender_totals = pivot.groupby("Ein_group")[metric].sum().sort_values(ascending=False)
anf_order = pivot.groupby("Anf_group")[metric].sum().sort_values(ascending=False).index.tolist()

# Build Mekko as stacked bars with variable widths
fig_mekko = go.Figure()
einsender_order = einsender_totals.index.tolist()
widths = einsender_totals.values
x_positions = np.concatenate([[0], np.cumsum(widths[:-1])])
colors = px.colors.qualitative.Set3 + px.colors.qualitative.Pastel

for i, anf in enumerate(anf_order):
    anf_vals = []
    for ein in einsender_order:
        val = pivot[(pivot["Ein_group"] == ein) & (pivot["Anf_group"] == anf)][metric].sum()
        total = einsender_totals[ein]
        anf_vals.append(val / total * 100 if total > 0 else 0)
    fig_mekko.add_trace(go.Bar(
        x=x_positions + widths / 2,
        y=anf_vals,
        width=widths * 0.95,
        name=anf[:40],
        marker_color=colors[i % len(colors)],
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            f"Anforderung: {anf[:40]}<br>"
            f"{metric}: %{{customdata[1]:,.0f}}<br>"
            "Anteil: %{y:.1f}%<extra></extra>"
        ),
        customdata=list(zip(
            einsender_order,
            [pivot[(pivot["Ein_group"] == ein) & (pivot["Anf_group"] == anf)][metric].sum() for ein in einsender_order],
        )),
    ))

fig_mekko.update_layout(
    barmode="stack",
    title=f"Mekko-Chart: Einsender × Anforderungsgruppen ({metric})",
    xaxis=dict(
        tickvals=x_positions + widths / 2,
        ticktext=[e[:20] for e in einsender_order],
        tickangle=-45,
        title="Einsender (Breite = Gesamtvolumen)",
    ),
    yaxis=dict(title="Anteil (%)", range=[0, 100]),
    height=600,
    legend=dict(orientation="h", yanchor="top", y=-0.35, xanchor="center", x=0.5),
)
st.plotly_chart(fig_mekko, use_container_width=True)

# ============================================================
# 5. HEATMAP: EINSENDER × ANFORDERUNG
# ============================================================
st.header("5 · Heatmap: Einsender × Anforderungen")

n_heat_ein = st.slider("Top-Einsender", 5, 25, 15, key="heat_ein")
n_heat_anf = st.slider("Top-Anforderungen", 5, 30, 15, key="heat_anf")

heat_einsender = df.groupby("Einsender")[metric].sum().nlargest(n_heat_ein).index
heat_anf = df.groupby("Anforderung")[metric].sum().nlargest(n_heat_anf).index

df_heat = df[df["Einsender"].isin(heat_einsender) & df["Anforderung"].isin(heat_anf)]
heat_pivot = df_heat.pivot_table(index="Anforderung", columns="Einsender", values=metric, aggfunc="sum", fill_value=0)
heat_pivot = heat_pivot.loc[heat_anf.intersection(heat_pivot.index), heat_einsender.intersection(heat_pivot.columns)]

fig_heat = px.imshow(
    heat_pivot.values,
    x=[c[:25] for c in heat_pivot.columns],
    y=[r[:35] for r in heat_pivot.index],
    color_continuous_scale="YlOrRd",
    aspect="auto",
    title=f"Heatmap ({metric})",
    labels=dict(color=metric),
)
fig_heat.update_layout(height=max(400, n_heat_anf * 28), xaxis_tickangle=-45)
st.plotly_chart(fig_heat, use_container_width=True)

# ============================================================
# 6. PARETO-ANALYSE
# ============================================================
st.header("6 · Pareto-Analyse")

pareto_level = st.radio("Pareto-Ebene", ["Einsender", "Anforderung"], horizontal=True)
pareto_data = df.groupby(pareto_level)[metric].sum().sort_values(ascending=False).reset_index()
pareto_data["Kumulativ %"] = pareto_data[metric].cumsum() / pareto_data[metric].sum() * 100

fig_pareto = go.Figure()
fig_pareto.add_trace(go.Bar(
    x=pareto_data[pareto_level], y=pareto_data[metric],
    name=metric, marker_color="steelblue",
))
fig_pareto.add_trace(go.Scatter(
    x=pareto_data[pareto_level], y=pareto_data["Kumulativ %"],
    name="Kumulativ %", yaxis="y2",
    line=dict(color="firebrick", width=2),
))
fig_pareto.add_hline(y=80, line_dash="dash", line_color="gray", yref="y2",
                     annotation_text="80 %", annotation_position="top left")
fig_pareto.update_layout(
    title=f"Pareto-Analyse — {pareto_level} ({metric})",
    yaxis=dict(title=metric),
    yaxis2=dict(title="Kumulativ %", overlaying="y", side="right", range=[0, 105]),
    xaxis_tickangle=-45, height=500, showlegend=True,
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
)
st.plotly_chart(fig_pareto, use_container_width=True)

# ============================================================
# 7. VERGLEICH ZWEIER EINSENDER
# ============================================================
st.header("7 · Vergleich zweier Einsender")

col_e, col_f = st.columns(2)
with col_e:
    ein_a = st.selectbox("Einsender A", einsender_list, index=0, key="cmp_a")
with col_f:
    ein_b = st.selectbox("Einsender B", einsender_list, index=min(1, len(einsender_list)-1), key="cmp_b")

df_a = df[df["Einsender"] == ein_a].set_index("Anforderung")[metric].rename(ein_a)
df_b = df[df["Einsender"] == ein_b].set_index("Anforderung")[metric].rename(ein_b)
cmp = pd.concat([df_a, df_b], axis=1).fillna(0)
cmp["Max"] = cmp.max(axis=1)
cmp = cmp.nlargest(15, "Max").drop(columns="Max")

fig_cmp = go.Figure()
fig_cmp.add_trace(go.Bar(y=cmp.index, x=cmp[ein_a], name=ein_a[:25], orientation="h", marker_color="steelblue"))
fig_cmp.add_trace(go.Bar(y=cmp.index, x=cmp[ein_b], name=ein_b[:25], orientation="h", marker_color="coral"))
fig_cmp.update_layout(
    barmode="group", title=f"Vergleich: Top 15 Anforderungen ({metric})",
    yaxis=dict(autorange="reversed"), height=500,
)
st.plotly_chart(fig_cmp, use_container_width=True)

# ============================================================
# 8. DATEN-EXPLORER
# ============================================================
st.header("8 · Daten-Explorer")

with st.expander("Rohdaten filtern & durchsuchen"):
    search = st.text_input("Freitextsuche (Einsender oder Anforderung)")
    df_display = df.copy()
    if search:
        mask = (
            df_display["Einsender"].str.contains(search, case=False, na=False)
            | df_display["Anforderung"].str.contains(search, case=False, na=False)
        )
        df_display = df_display[mask]
    st.dataframe(df_display.sort_values(metric, ascending=False), use_container_width=True, height=500)
    st.download_button(
        "CSV herunterladen",
        df_display.to_csv(index=False).encode("utf-8"),
        "labor_daten_gefiltert.csv",
        "text/csv",
    )
