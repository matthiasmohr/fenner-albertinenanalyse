import os
import json
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from dotenv import load_dotenv
import streamlit_authenticator as stauth
from streamlit_authenticator.utilities import LoginError

load_dotenv()

st.set_page_config(page_title="Labor-Analyse", layout="wide")

# --- Multi-user authentication ---
LOGIN_LOG_PATH = "login_log.json"

def build_credentials():
    users = os.environ.get("AUTH_USERS", "").split(",")
    passwords = os.environ.get("AUTH_PASSWORDS", "").split(",")
    names = os.environ.get("AUTH_NAMES", "").split(",")
    roles = os.environ.get("AUTH_ROLES", "").split(",")
    credentials = {"usernames": {}}
    for user, pw, name, role in zip(users, passwords, names, roles):
        credentials["usernames"][user.strip()] = {
            "name": name.strip(),
            "password": pw.strip(),
            "role": role.strip(),
            "email": "",
        }
    return credentials

def _read_login_log():
    if not os.path.exists(LOGIN_LOG_PATH):
        return {}
    try:
        with open(LOGIN_LOG_PATH, "r") as f:
            content = f.read().strip()
            return json.loads(content) if content else {}
    except (json.JSONDecodeError, OSError):
        return {}

def log_login(username):
    log = _read_login_log()
    log[username] = datetime.now().isoformat(timespec="seconds")
    with open(LOGIN_LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)

credentials = build_credentials()
cookie_key = os.environ.get("AUTH_COOKIE_KEY", "default_secret_key")

authenticator = stauth.Authenticate(
    credentials,
    cookie_name="labor_analyse_auth",
    cookie_key=cookie_key,
    cookie_expiry_days=30,
    auto_hash=True,
)

LOGIN_FIELDS = {"Form name": "Anmeldung", "Username": "Benutzername", "Password": "Passwort", "Login": "Anmelden"}

try:
    authenticator.login(fields=LOGIN_FIELDS)
except LoginError:
    # Stale Cookie (z.B. Nutzer ist nicht mehr in AUTH_USERS): Cookie verwerfen,
    # logout-Flag setzen, damit get_cookie() den Token-Pfad überspringt, und Formular rendern.
    try:
        authenticator.cookie_controller.delete_cookie()
    except Exception:
        pass
    st.session_state["logout"] = True
    for key in ("authentication_status", "name", "username", "email", "roles", "login_logged"):
        st.session_state.pop(key, None)
    authenticator.login(fields=LOGIN_FIELDS)

if st.session_state.get("authentication_status") is None:
    st.stop()
elif st.session_state.get("authentication_status") is False:
    st.error("Benutzername oder Passwort falsch.")
    st.stop()

# --- Authenticated: log login ---
if not st.session_state.get("login_logged"):
    log_login(st.session_state["username"])
    st.session_state["login_logged"] = True

current_user = st.session_state["username"]
current_role = credentials["usernames"].get(current_user, {}).get("role", "user")

# --- Header info bar from env ---
betrachtungszeitraum = os.environ.get("BETRACHTUNGSZEITRAUM", "–")
datenimport_zeitpunkt = os.environ.get("DATENIMPORT_ZEITPUNKT", "–")
exclude_zero_goae_default = os.environ.get("EXCLUDE_ZERO_GOAE", "true").lower() == "true"

st.markdown(
    f"**Betrachtungszeitraum:** {betrachtungszeitraum} · "
    f"**Datenimport:** {datenimport_zeitpunkt} · "
    f"**Angemeldet als:** {credentials['usernames'].get(current_user, {}).get('name', current_user)}"
)
st.markdown("---")

# --- Admin: Login-Übersicht ---
if current_role == "admin":
    with st.expander("Admin: Letzte Logins"):
        log = _read_login_log()
        if log:
            login_data = []
            for username, timestamp in log.items():
                name = credentials["usernames"].get(username, {}).get("name", username)
                login_data.append({"Benutzer": name, "Benutzername": username, "Letzter Login": timestamp})
            st.dataframe(pd.DataFrame(login_data), use_container_width=True, hide_index=True)
        else:
            st.info("Noch keine Logins aufgezeichnet.")

# --- Labor selection ---
@st.cache_data
def find_labs():
    files = sorted(f for f in os.listdir("input") if f.endswith(".xlsx"))
    return {os.path.splitext(f)[0]: os.path.join("input", f) for f in files}

REQUIRED_COLUMNS = ["Einsender", "Analyse/Leistung", "Punktsumme", "Anzahl"]

@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Fehlende Pflicht-Spalten: {missing}\n\n"
            f"Erwartet werden genau diese Spalten: {REQUIRED_COLUMNS}"
        )
    return df[REQUIRED_COLUMNS]

labs = find_labs()
if not labs:
    st.error("Keine Excel-Dateien im Ordner `input/` gefunden.")
    st.stop()

selected_lab = st.sidebar.selectbox("Labor", list(labs.keys()))
try:
    df = load_data(labs[selected_lab])
except ValueError as e:
    st.error(str(e))
    st.stop()

st.title(f"Labor-Anforderungsanalyse — {selected_lab}")

# --- Global filters ---
metric = st.sidebar.radio("Kennzahl", ["Punktsumme", "Anzahl"], index=0)
exclude_zero_goae = st.sidebar.checkbox(
    "Ohne Punktsumme ausschließen",
    value=exclude_zero_goae_default,
    help="Einträge mit Punktsumme = 0 ausblenden (i.d.R. interne Steuerkennzeichen und Statistiken)",
)
if exclude_zero_goae:
    df = df[df["Punktsumme"] > 0]
st.sidebar.markdown("---")
st.sidebar.markdown(f"**{df['Einsender'].nunique()}** Einsender · **{df['Analyse/Leistung'].nunique()}** Analysen/Leistungen")
st.sidebar.markdown(f"**{df['Punktsumme'].sum():,.0f}** Punktsumme · **{df['Anzahl'].sum():,.0f}** Anzahl")

# ============================================================
# 1. ÜBERBLICK
# ============================================================
st.header("1 · Überblick")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Einsender", df["Einsender"].nunique())
col2.metric("Analysen/Leistungen", df["Analyse/Leistung"].nunique())
col3.metric("Σ Punktsumme", f"{df['Punktsumme'].sum():,.0f}")
col4.metric("Σ Anzahl", f"{df['Anzahl'].sum():,.0f}")

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
st.header("2 · Top 30 Analysen/Leistungen je Einsender")

einsender_list = einsender_agg["Einsender"].tolist()
selected_einsender = st.selectbox("Einsender auswählen", einsender_list, key="top30_einsender")

df_filtered = df[df["Einsender"] == selected_einsender]
top30_anf = df_filtered.nlargest(30, metric)

col_a, col_b = st.columns([2, 1])
with col_a:
    fig_top30_anf = px.bar(
        top30_anf, x=metric, y="Analyse/Leistung", orientation="h",
        title=f"Top 30 Analysen/Leistungen — {selected_einsender}",
        color=metric, color_continuous_scale="Teal",
    )
    fig_top30_anf.update_layout(yaxis=dict(autorange="reversed"), height=max(500, len(top30_anf) * 22), showlegend=False)
    st.plotly_chart(fig_top30_anf, use_container_width=True)
with col_b:
    st.dataframe(
        top30_anf[["Analyse/Leistung", "Punktsumme", "Anzahl"]].reset_index(drop=True),
        use_container_width=True, height=max(500, len(top30_anf) * 22),
    )

# ============================================================
# 3. TOP 30 EINSENDER JE ANFORDERUNG
# ============================================================
st.header("3 · Top 30 Einsender je Analyse/Leistung")

anforderung_agg = df.groupby("Analyse/Leistung")[metric].sum().sort_values(ascending=False)
anforderung_list = anforderung_agg.index.tolist()
selected_anforderung = st.selectbox("Analyse/Leistung auswählen", anforderung_list, key="top30_anforderung")

df_anf = df[df["Analyse/Leistung"] == selected_anforderung].nlargest(30, metric)

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
        df_anf[["Einsender", "Punktsumme", "Anzahl"]].reset_index(drop=True),
        use_container_width=True, height=max(500, len(df_anf) * 22),
    )

# ============================================================
# 4. MEKKO-CHART
# ============================================================
st.header("4 · Mekko-Chart (Einsender × Analyse/Leistungsgruppen)")

n_top_einsender = st.slider("Anzahl Top-Einsender", 5, 20, 10, key="mekko_n")
n_top_anf = st.slider("Anzahl Top-Analyse/Leistungsgruppen", 5, 15, 8, key="mekko_anf")

top_einsender = df.groupby("Einsender")[metric].sum().nlargest(n_top_einsender).index.tolist()
top_anforderungen = df.groupby("Analyse/Leistung")[metric].sum().nlargest(n_top_anf).index.tolist()

df_mekko = df.copy()
df_mekko["Anf_group"] = df_mekko["Analyse/Leistung"].where(df_mekko["Analyse/Leistung"].isin(top_anforderungen), "Sonstige")
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
            f"Analyse/Leistung: {anf[:40]}<br>"
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
    title=f"Mekko-Chart: Einsender × Analyse/Leistungsgruppen ({metric})",
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
st.header("5 · Heatmap: Einsender × Analysen/Leistungen")

n_heat_ein = st.slider("Top-Einsender", 5, 25, 15, key="heat_ein")
n_heat_anf = st.slider("Top-Analysen/Leistungen", 5, 30, 15, key="heat_anf")

heat_einsender = df.groupby("Einsender")[metric].sum().nlargest(n_heat_ein).index
heat_anf = df.groupby("Analyse/Leistung")[metric].sum().nlargest(n_heat_anf).index

df_heat = df[df["Einsender"].isin(heat_einsender) & df["Analyse/Leistung"].isin(heat_anf)]
heat_pivot = df_heat.pivot_table(index="Analyse/Leistung", columns="Einsender", values=metric, aggfunc="sum", fill_value=0)
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

pareto_level = st.radio("Pareto-Ebene", ["Einsender", "Analyse/Leistung"], horizontal=True)
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

if ein_a == ein_b:
    st.warning("Bitte zwei verschiedene Einsender auswählen.")
    st.stop()

df_a = df[df["Einsender"] == ein_a].set_index("Analyse/Leistung")[metric].rename(ein_a)
df_b = df[df["Einsender"] == ein_b].set_index("Analyse/Leistung")[metric].rename(ein_b)
cmp = pd.concat([df_a, df_b], axis=1).fillna(0)
cmp["Max"] = cmp.max(axis=1)
cmp = cmp.nlargest(15, "Max").drop(columns="Max")

fig_cmp = go.Figure()
fig_cmp.add_trace(go.Bar(y=cmp.index, x=cmp[ein_a], name=ein_a[:25], orientation="h", marker_color="steelblue"))
fig_cmp.add_trace(go.Bar(y=cmp.index, x=cmp[ein_b], name=ein_b[:25], orientation="h", marker_color="coral"))
fig_cmp.update_layout(
    barmode="group", title=f"Vergleich: Top 15 Analysen/Leistungen ({metric})",
    yaxis=dict(autorange="reversed"), height=500,
)
st.plotly_chart(fig_cmp, use_container_width=True)

# ============================================================
# 8. DATEN-EXPLORER
# ============================================================
st.header("8 · Daten-Explorer")

with st.expander("Rohdaten filtern & durchsuchen"):
    search = st.text_input("Freitextsuche (Einsender oder Analyse/Leistung)")
    df_display = df.copy()
    if search:
        mask = (
            df_display["Einsender"].str.contains(search, case=False, na=False)
            | df_display["Analyse/Leistung"].str.contains(search, case=False, na=False)
        )
        df_display = df_display[mask]
    st.dataframe(df_display.sort_values(metric, ascending=False), use_container_width=True, height=500)
    st.download_button(
        "CSV herunterladen",
        df_display.to_csv(index=False).encode("utf-8"),
        "labor_daten_gefiltert.csv",
        "text/csv",
    )
