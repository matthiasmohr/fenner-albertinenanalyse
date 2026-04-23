import os
import re
import json
import unicodedata
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
show_kosten = os.environ.get("SHOW_KOSTEN", "true").lower() == "true"

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
YEAR_SUFFIX_RE = re.compile(r"-(\d{4})$")

@st.cache_data
def find_labs():
    """Gruppiert Excel-Dateien in `input/` nach Basisname und erkennt Jahres-Suffix `-YYYY`.

    Rückgabe: {base_name: {year_or_None: path}}. Labore ohne Jahres-Suffix bekommen `None`
    als Key, sodass die bisherige Einjahres-Logik unverändert greifen kann.
    """
    groups: dict[str, dict] = {}
    # macOS liefert Dateinamen oft in NFD (zerlegte Umlaute) zurück, andere Quellen
    # in NFC. Ohne Normalisierung bricht das Gruppieren, weil "Präsenzlabor..." in
    # beiden Formen als unterschiedliche Keys auftaucht.
    for f in sorted(os.listdir("input"), key=lambda s: unicodedata.normalize("NFC", s)):
        if not f.endswith(".xlsx"):
            continue
        name = unicodedata.normalize("NFC", os.path.splitext(f)[0])
        m = YEAR_SUFFIX_RE.search(name)
        if m:
            base = name[: m.start()]
            year = int(m.group(1))
        else:
            base = name
            year = None
        groups.setdefault(base, {})[year] = os.path.join("input", f)

    # Dateien ohne Jahres-Suffix gelten als "aktuelles Jahr". Wenn die Gruppe auch
    # ein Jahr mit Suffix hat (z.B. -2024), wird die suffix-lose Datei als
    # (max_year + 1) interpretiert, damit der Jahresvergleich funktioniert.
    current_year_env = os.environ.get("CURRENT_YEAR")
    try:
        current_year_default = int(current_year_env) if current_year_env else None
    except ValueError:
        current_year_default = None

    for base, files in list(groups.items()):
        if None not in files:
            continue
        yeared = [y for y in files.keys() if y is not None]
        if yeared:
            inferred = current_year_default or (max(yeared) + 1)
            files[inferred] = files.pop(None)
    return groups

REQUIRED_COLUMNS = ["Einsender", "Analyse/Leistung", "Punktsumme", "Anzahl"]

@st.cache_data
def load_data(paths_by_year):
    """Lädt eine oder mehrere Jahres-Dateien und fügt eine `Jahr`-Spalte hinzu, falls
    mindestens eine Datei ein Jahr mitbringt. Bei einzelnem Jahr-losen Lab bleibt das
    DataFrame unverändert (ohne `Jahr`-Spalte)."""
    has_year = any(y is not None for y in paths_by_year.keys())
    dfs = []
    for year, path in paths_by_year.items():
        d = pd.read_excel(path)
        missing = [c for c in REQUIRED_COLUMNS if c not in d.columns]
        if missing:
            raise ValueError(
                f"Fehlende Pflicht-Spalten in {os.path.basename(path)}: {missing}\n\n"
                f"Erwartet werden mindestens diese Spalten: {REQUIRED_COLUMNS}"
            )
        d = d[REQUIRED_COLUMNS].copy()
        if has_year:
            d["Jahr"] = str(year) if year is not None else "—"
        dfs.append(d)
    return pd.concat(dfs, ignore_index=True)

labs = find_labs()
if not labs:
    st.error("Keine Excel-Dateien im Ordner `input/` gefunden.")
    st.stop()

selected_lab = st.sidebar.selectbox("Labor", list(labs.keys()))
paths_by_year = labs[selected_lab]
years_available = sorted([y for y in paths_by_year.keys() if y is not None])
has_years = len(years_available) > 0

# Jahres-Auswahl: bei >=2 verfügbaren Jahren Vergleichsmodus anbieten
year_mode = None  # "Beide" oder str(year) oder None (Lab ohne Jahreszuordnung)
if len(years_available) >= 2:
    year_options = ["Beide Jahre (Vergleich)"] + [str(y) for y in years_available]
    year_mode = st.sidebar.radio("Jahr", year_options, index=0)
elif len(years_available) == 1:
    year_mode = str(years_available[0])

try:
    df = load_data(paths_by_year)
except ValueError as e:
    st.error(str(e))
    st.stop()

# Vergleichsmodus nur, wenn mehrere Jahre geladen sind und Nutzer "Beide" wählt
compare_mode = year_mode is not None and year_mode.startswith("Beide")

# Bei Einzeljahr-Auswahl auf dieses Jahr filtern
if has_years and not compare_mode:
    df = df[df["Jahr"] == year_mode]

YEAR_COLORS = {"2024": "#4C78A8", "2025": "#F58518", "2026": "#54A24B"}

st.title(f"Labor-Anforderungsanalyse — {selected_lab}")
if compare_mode:
    st.caption(f"Vergleichsmodus: {' vs. '.join(str(y) for y in years_available)}")
elif has_years:
    st.caption(f"Jahr: {year_mode}")

# --- Global filters ---
GOAE_BETRAG = 0.0582873  # GOÄ-Punktwert einfacher Satz (§5 Abs. 1 GOÄ)
_metric_options = ["Punktsumme", "Anzahl"] + (["Kalk. Kosten"] if show_kosten else [])
metric = st.sidebar.radio("Kennzahl", _metric_options, index=0)
if show_kosten:
    goae_faktor = st.sidebar.number_input(
        "GOÄ-Faktor",
        min_value=0.01, max_value=10.0, value=0.40, step=0.01, format="%.2f",
        help=f"Kalk. Kosten = Punktsumme × {GOAE_BETRAG:.7f} € × GOÄ-Faktor",
        disabled=(metric != "Kalk. Kosten"),
    )
else:
    goae_faktor = 0.40
exclude_zero_goae = st.sidebar.checkbox(
    "Ohne Punktsumme ausschließen",
    value=exclude_zero_goae_default,
    help="Einträge mit Punktsumme = 0 ausblenden (i.d.R. interne Steuerkennzeichen und Statistiken)",
)
if exclude_zero_goae:
    df = df[df["Punktsumme"] > 0]
df["Kalk. Kosten"] = df["Punktsumme"] * GOAE_BETRAG * goae_faktor
st.sidebar.markdown("---")
st.sidebar.markdown(f"**{df['Einsender'].nunique()}** Einsender · **{df['Analyse/Leistung'].nunique()}** Analysen/Leistungen")
st.sidebar.markdown(f"**{df['Punktsumme'].sum():,.0f}** Punktsumme · **{df['Anzahl'].sum():,.0f}** Anzahl")
if show_kosten:
    st.sidebar.markdown(f"**{df['Kalk. Kosten'].sum():,.2f} €** Kalk. Kosten")

# Metric formatting helpers
_cost_mode = metric == "Kalk. Kosten"
_tick_fmt = dict(tickformat=",.2f", tickprefix="€\u202f") if _cost_mode else {}
_cdata_fmt = "€\u202f%{customdata[1]:,.2f}" if _cost_mode else "%{customdata[1]:,.0f}"
_fmt_metric = (lambda v: f"{v:,.2f} €") if _cost_mode else (lambda v: f"{v:,.0f}")

def _delta_pct(curr, prev):
    if prev is None or prev == 0:
        return None
    return (curr - prev) / prev * 100

# ============================================================
# 1. ÜBERBLICK
# ============================================================
st.header("1 · Überblick")

if compare_mode:
    years_sorted = sorted(df["Jahr"].unique())
    # Pro Jahr eine Zeile mit KPIs + Delta ggü. Vorjahr
    st.markdown("**Kennzahlen je Jahr**")
    for i, yr in enumerate(years_sorted):
        d_y = df[df["Jahr"] == yr]
        prev = df[df["Jahr"] == years_sorted[i - 1]] if i > 0 else None
        ncols = 5 if show_kosten else 4
        cols = st.columns(ncols)
        cols[0].metric(f"Einsender {yr}", d_y["Einsender"].nunique())
        cols[1].metric(f"Analysen/Leistungen {yr}", d_y["Analyse/Leistung"].nunique())
        d_pkt = _delta_pct(d_y["Punktsumme"].sum(), prev["Punktsumme"].sum() if prev is not None else None)
        cols[2].metric(f"Σ Punktsumme {yr}", f"{d_y['Punktsumme'].sum():,.0f}",
                       delta=(f"{d_pkt:+.1f} %" if d_pkt is not None else None))
        d_anz = _delta_pct(d_y["Anzahl"].sum(), prev["Anzahl"].sum() if prev is not None else None)
        cols[3].metric(f"Σ Anzahl {yr}", f"{d_y['Anzahl'].sum():,.0f}",
                       delta=(f"{d_anz:+.1f} %" if d_anz is not None else None))
        if show_kosten:
            d_kost = _delta_pct(d_y["Kalk. Kosten"].sum(), prev["Kalk. Kosten"].sum() if prev is not None else None)
            cols[4].metric(f"Σ Kalk. Kosten {yr}", f"{d_y['Kalk. Kosten'].sum():,.2f} €",
                           delta=(f"{d_kost:+.1f} %" if d_kost is not None else None))
else:
    if show_kosten:
        col1, col2, col3, col4, col5 = st.columns(5)
        col5.metric("Σ Kalk. Kosten", f"{df['Kalk. Kosten'].sum():,.2f} €")
    else:
        col1, col2, col3, col4 = st.columns(4)
    col1.metric("Einsender", df["Einsender"].nunique())
    col2.metric("Analysen/Leistungen", df["Analyse/Leistung"].nunique())
    col3.metric("Σ Punktsumme", f"{df['Punktsumme'].sum():,.0f}")
    col4.metric("Σ Anzahl", f"{df['Anzahl'].sum():,.0f}")

# Aggregationen für Sortier-Reihenfolge (über alle Jahre hinweg)
einsender_totals = df.groupby("Einsender")[metric].sum().sort_values(ascending=False)
analyten_totals = df.groupby("Analyse/Leistung")[metric].sum().sort_values(ascending=False)

tab_ein, tab_ana = st.tabs(["Top Einsender", "Top Analyten"])

with tab_ein:
    if compare_mode:
        ein_by_year = (
            df.groupby(["Einsender", "Jahr"])[metric].sum().reset_index()
        )
        ein_by_year["Einsender"] = pd.Categorical(
            ein_by_year["Einsender"], categories=einsender_totals.index.tolist(), ordered=True
        )
        ein_by_year = ein_by_year.sort_values(["Einsender", "Jahr"])
        fig_bar = px.bar(
            ein_by_year, x="Einsender", y=metric, color="Jahr", barmode="group",
            title=f"{metric} je Einsender — Jahresvergleich",
            color_discrete_map=YEAR_COLORS,
        )
    else:
        einsender_agg = einsender_totals.reset_index()
        fig_bar = px.bar(
            einsender_agg, x="Einsender", y=metric,
            title=f"{metric} je Einsender (alle)",
            color=metric, color_continuous_scale="Blues",
        )
    fig_bar.update_layout(xaxis_tickangle=-45, height=500,
                          showlegend=compare_mode,
                          yaxis=_tick_fmt,
                          coloraxis_colorbar=dict(**_tick_fmt) if not compare_mode else None)
    st.plotly_chart(fig_bar, use_container_width=True)

with tab_ana:
    n_top_ana = st.slider("Anzahl Top-Analyten", 10, 50, 25, key="ueberblick_top_ana")
    top_ana_names = analyten_totals.head(n_top_ana).index.tolist()
    if compare_mode:
        ana_by_year = (
            df[df["Analyse/Leistung"].isin(top_ana_names)]
            .groupby(["Analyse/Leistung", "Jahr"])[metric].sum().reset_index()
        )
        ana_by_year["Analyse/Leistung"] = pd.Categorical(
            ana_by_year["Analyse/Leistung"], categories=top_ana_names, ordered=True
        )
        ana_by_year = ana_by_year.sort_values(["Analyse/Leistung", "Jahr"])
        fig_ana = px.bar(
            ana_by_year, x=metric, y="Analyse/Leistung", color="Jahr", orientation="h", barmode="group",
            title=f"Top {n_top_ana} Analyten/Leistungen — Jahresvergleich",
            color_discrete_map=YEAR_COLORS,
        )
    else:
        top_ana = analyten_totals.head(n_top_ana).reset_index()
        fig_ana = px.bar(
            top_ana, x=metric, y="Analyse/Leistung", orientation="h",
            title=f"Top {n_top_ana} Analyten/Leistungen — {metric} (global)",
            color=metric, color_continuous_scale="Teal",
        )
    fig_ana.update_layout(yaxis=dict(autorange="reversed"), height=max(500, n_top_ana * 22),
                          showlegend=compare_mode,
                          xaxis=_tick_fmt,
                          coloraxis_colorbar=dict(**_tick_fmt) if not compare_mode else None)
    st.plotly_chart(fig_ana, use_container_width=True)

# ============================================================
# 2. TOP 30 ANFORDERUNGEN JE EINSENDER
# ============================================================
st.header("2 · Top 30 Analysen/Leistungen je Einsender")

einsender_list = einsender_totals.index.tolist()
selected_einsender = st.selectbox("Einsender auswählen", einsender_list, key="top30_einsender")

df_filtered = df[df["Einsender"] == selected_einsender]
# Top 30 Analysen anhand Summe über alle Jahre wählen
top30_names = (
    df_filtered.groupby("Analyse/Leistung")[metric].sum().nlargest(30).index.tolist()
)
top30_anf = df_filtered[df_filtered["Analyse/Leistung"].isin(top30_names)]

col_a, col_b = st.columns([2, 1])
with col_a:
    if compare_mode:
        top30_plot = (
            top30_anf.groupby(["Analyse/Leistung", "Jahr"])[metric].sum().reset_index()
        )
        top30_plot["Analyse/Leistung"] = pd.Categorical(
            top30_plot["Analyse/Leistung"], categories=top30_names, ordered=True
        )
        top30_plot = top30_plot.sort_values(["Analyse/Leistung", "Jahr"])
        fig_top30_anf = px.bar(
            top30_plot, x=metric, y="Analyse/Leistung", color="Jahr", orientation="h", barmode="group",
            title=f"Top 30 Analysen/Leistungen — {selected_einsender} (Jahresvergleich)",
            color_discrete_map=YEAR_COLORS,
        )
    else:
        top30_plot = top30_anf.nlargest(30, metric)
        fig_top30_anf = px.bar(
            top30_plot, x=metric, y="Analyse/Leistung", orientation="h",
            title=f"Top 30 Analysen/Leistungen — {selected_einsender}",
            color=metric, color_continuous_scale="Teal",
        )
    fig_top30_anf.update_layout(yaxis=dict(autorange="reversed"), height=max(500, len(top30_names) * 22),
                                showlegend=compare_mode,
                                xaxis=_tick_fmt,
                                coloraxis_colorbar=dict(**_tick_fmt) if not compare_mode else None)
    st.plotly_chart(fig_top30_anf, use_container_width=True)
with col_b:
    tbl_cols = ["Analyse/Leistung"] + (["Jahr"] if compare_mode else []) + ["Punktsumme", "Anzahl"] + (["Kalk. Kosten"] if show_kosten else [])
    tbl = (
        top30_anf.sort_values(metric, ascending=False)[tbl_cols].reset_index(drop=True)
        if not compare_mode else
        top30_anf.sort_values(["Analyse/Leistung", "Jahr"])[tbl_cols].reset_index(drop=True)
    )
    st.dataframe(tbl, use_container_width=True, height=max(500, len(top30_anf) * 22))

# ============================================================
# 3. TOP 30 EINSENDER JE ANFORDERUNG
# ============================================================
st.header("3 · Top 30 Einsender je Analyse/Leistung")

anforderung_list = analyten_totals.index.tolist()
selected_anforderung = st.selectbox("Analyse/Leistung auswählen", anforderung_list, key="top30_anforderung")

df_anf_full = df[df["Analyse/Leistung"] == selected_anforderung]
top30_ein_names = (
    df_anf_full.groupby("Einsender")[metric].sum().nlargest(30).index.tolist()
)
df_anf = df_anf_full[df_anf_full["Einsender"].isin(top30_ein_names)]

col_c, col_d = st.columns([2, 1])
with col_c:
    if compare_mode:
        anf_plot = df_anf.groupby(["Einsender", "Jahr"])[metric].sum().reset_index()
        anf_plot["Einsender"] = pd.Categorical(anf_plot["Einsender"], categories=top30_ein_names, ordered=True)
        anf_plot = anf_plot.sort_values(["Einsender", "Jahr"])
        fig_top30_ein = px.bar(
            anf_plot, x=metric, y="Einsender", color="Jahr", orientation="h", barmode="group",
            title=f"Top 30 Einsender — {selected_anforderung} (Jahresvergleich)",
            color_discrete_map=YEAR_COLORS,
        )
    else:
        anf_plot = df_anf.nlargest(30, metric)
        fig_top30_ein = px.bar(
            anf_plot, x=metric, y="Einsender", orientation="h",
            title=f"Top 30 Einsender — {selected_anforderung}",
            color=metric, color_continuous_scale="Oranges",
        )
    fig_top30_ein.update_layout(yaxis=dict(autorange="reversed"), height=max(500, len(top30_ein_names) * 22),
                                showlegend=compare_mode,
                                xaxis=_tick_fmt,
                                coloraxis_colorbar=dict(**_tick_fmt) if not compare_mode else None)
    st.plotly_chart(fig_top30_ein, use_container_width=True)
with col_d:
    tbl_cols = ["Einsender"] + (["Jahr"] if compare_mode else []) + ["Punktsumme", "Anzahl"] + (["Kalk. Kosten"] if show_kosten else [])
    tbl = (
        df_anf.sort_values(metric, ascending=False)[tbl_cols].reset_index(drop=True)
        if not compare_mode else
        df_anf.sort_values(["Einsender", "Jahr"])[tbl_cols].reset_index(drop=True)
    )
    st.dataframe(tbl, use_container_width=True, height=max(500, len(df_anf) * 22))

# ============================================================
# 4. MEKKO-CHART
# ============================================================
st.header("4 · Mekko-Chart (Einsender × Analyse/Leistungsgruppen)")

# Mekko ist für einen Vergleich zweier Jahre visuell überladen. Im Vergleichsmodus
# lässt der Nutzer das darzustellende Jahr wählen.
if compare_mode:
    mekko_year = st.radio("Jahr für Mekko", sorted(df["Jahr"].unique()), horizontal=True, key="mekko_year")
    df_for_mekko = df[df["Jahr"] == mekko_year]
else:
    df_for_mekko = df

n_top_einsender = st.slider("Anzahl Top-Einsender", 5, 20, 10, key="mekko_n")
n_top_anf = st.slider("Anzahl Top-Analyse/Leistungsgruppen", 5, 15, 8, key="mekko_anf")

top_einsender = df_for_mekko.groupby("Einsender")[metric].sum().nlargest(n_top_einsender).index.tolist()
top_anforderungen = df_for_mekko.groupby("Analyse/Leistung")[metric].sum().nlargest(n_top_anf).index.tolist()

df_mekko = df_for_mekko.copy()
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
            f"{metric}: {_cdata_fmt}<br>"
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

if compare_mode:
    years_sorted = sorted(df["Jahr"].unique())
    heat_mode = st.radio(
        "Darstellung",
        [f"Differenz ({years_sorted[-1]} − {years_sorted[0]})"] + years_sorted,
        horizontal=True,
        key="heat_mode",
    )
else:
    heat_mode = None

n_heat_ein = st.slider("Top-Einsender", 5, 25, 15, key="heat_ein")
n_heat_anf = st.slider("Top-Analysen/Leistungen", 5, 30, 15, key="heat_anf")

# Top-Auswahl aus Gesamtdaten (Jahre addiert), damit die Achsen stabil sind
heat_einsender = df.groupby("Einsender")[metric].sum().nlargest(n_heat_ein).index
heat_anf = df.groupby("Analyse/Leistung")[metric].sum().nlargest(n_heat_anf).index

if compare_mode and heat_mode and heat_mode.startswith("Differenz"):
    ya, yb = years_sorted[0], years_sorted[-1]
    p_a = (
        df[(df["Jahr"] == ya) & df["Einsender"].isin(heat_einsender) & df["Analyse/Leistung"].isin(heat_anf)]
        .pivot_table(index="Analyse/Leistung", columns="Einsender", values=metric, aggfunc="sum", fill_value=0)
    )
    p_b = (
        df[(df["Jahr"] == yb) & df["Einsender"].isin(heat_einsender) & df["Analyse/Leistung"].isin(heat_anf)]
        .pivot_table(index="Analyse/Leistung", columns="Einsender", values=metric, aggfunc="sum", fill_value=0)
    )
    # Auf gemeinsame Achse ausrichten
    p_a = p_a.reindex(index=heat_anf, columns=heat_einsender, fill_value=0)
    p_b = p_b.reindex(index=heat_anf, columns=heat_einsender, fill_value=0)
    heat_pivot = p_b - p_a
    vmax = float(np.nanmax(np.abs(heat_pivot.values))) if heat_pivot.size else 0.0
    fig_heat = px.imshow(
        heat_pivot.values,
        x=[c[:25] for c in heat_pivot.columns],
        y=[r[:35] for r in heat_pivot.index],
        color_continuous_scale="RdBu_r",
        zmin=-vmax, zmax=vmax,
        aspect="auto",
        title=f"Heatmap Differenz {yb} − {ya} ({metric}) · rot = Wachstum, blau = Rückgang",
        labels=dict(color=f"Δ {metric}"),
    )
else:
    year_filter = heat_mode if (compare_mode and heat_mode in years_sorted) else None
    df_heat_src = df if year_filter is None else df[df["Jahr"] == year_filter]
    df_heat = df_heat_src[df_heat_src["Einsender"].isin(heat_einsender) & df_heat_src["Analyse/Leistung"].isin(heat_anf)]
    heat_pivot = df_heat.pivot_table(index="Analyse/Leistung", columns="Einsender", values=metric, aggfunc="sum", fill_value=0)
    heat_pivot = heat_pivot.reindex(index=heat_anf, columns=heat_einsender, fill_value=0)
    title_suffix = f" · {year_filter}" if year_filter else ""
    fig_heat = px.imshow(
        heat_pivot.values,
        x=[c[:25] for c in heat_pivot.columns],
        y=[r[:35] for r in heat_pivot.index],
        color_continuous_scale="YlOrRd",
        aspect="auto",
        title=f"Heatmap ({metric}){title_suffix}",
        labels=dict(color=metric),
    )
fig_heat.update_layout(height=max(400, n_heat_anf * 28), xaxis_tickangle=-45,
                       coloraxis_colorbar=dict(**_tick_fmt))
st.plotly_chart(fig_heat, use_container_width=True)

# ============================================================
# 6. PARETO-ANALYSE
# ============================================================
st.header("6 · Pareto-Analyse")

pareto_level = st.radio("Pareto-Ebene", ["Einsender", "Analyse/Leistung"], horizontal=True)

fig_pareto = go.Figure()
if compare_mode:
    years_sorted = sorted(df["Jahr"].unique())
    # Einheitliche X-Achsen-Reihenfolge: Top nach Gesamtsumme über alle Jahre
    order = df.groupby(pareto_level)[metric].sum().sort_values(ascending=False).index.tolist()
    for yr in years_sorted:
        d_y = df[df["Jahr"] == yr].groupby(pareto_level)[metric].sum()
        d_y = d_y.reindex(order, fill_value=0)
        cum = d_y.cumsum() / d_y.sum() * 100 if d_y.sum() > 0 else d_y * 0
        color = YEAR_COLORS.get(yr, None)
        fig_pareto.add_trace(go.Bar(
            x=order, y=d_y.values, name=f"{metric} {yr}",
            marker_color=color, opacity=0.85,
        ))
        fig_pareto.add_trace(go.Scatter(
            x=order, y=cum.values, name=f"Kumulativ % {yr}",
            yaxis="y2", line=dict(color=color, width=2, dash="dot"),
        ))
    fig_pareto.update_layout(barmode="group")
else:
    pareto_data = df.groupby(pareto_level)[metric].sum().sort_values(ascending=False).reset_index()
    pareto_data["Kumulativ %"] = pareto_data[metric].cumsum() / pareto_data[metric].sum() * 100
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
    title=f"Pareto-Analyse — {pareto_level} ({metric})"
          + (" · Jahresvergleich" if compare_mode else ""),
    yaxis=dict(title=metric, **_tick_fmt),
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

# Für Vergleich zweier Einsender: bei Beide-Mode ein konkretes Jahr wählen lassen,
# damit der Chart lesbar bleibt. Gruppierte Bars pro Jahr × Einsender würden zu dicht.
if compare_mode:
    cmp_year = st.radio("Jahr für diesen Vergleich", sorted(df["Jahr"].unique()), horizontal=True, key="cmp_year")
    df_cmp = df[df["Jahr"] == cmp_year]
else:
    cmp_year = None
    df_cmp = df

df_a = df_cmp[df_cmp["Einsender"] == ein_a].groupby("Analyse/Leistung")[metric].sum().rename(ein_a)
df_b = df_cmp[df_cmp["Einsender"] == ein_b].groupby("Analyse/Leistung")[metric].sum().rename(ein_b)
cmp = pd.concat([df_a, df_b], axis=1).fillna(0)
cmp["Max"] = cmp.max(axis=1)
cmp = cmp.nlargest(15, "Max").drop(columns="Max")

fig_cmp = go.Figure()
fig_cmp.add_trace(go.Bar(y=cmp.index, x=cmp[ein_a], name=ein_a[:25], orientation="h", marker_color="steelblue"))
fig_cmp.add_trace(go.Bar(y=cmp.index, x=cmp[ein_b], name=ein_b[:25], orientation="h", marker_color="coral"))
fig_cmp.update_layout(
    barmode="group",
    title=f"Vergleich: Top 15 Analysen/Leistungen ({metric})"
          + (f" · {cmp_year}" if cmp_year else ""),
    yaxis=dict(autorange="reversed"), height=500,
    xaxis=_tick_fmt,
)
st.plotly_chart(fig_cmp, use_container_width=True)

# ============================================================
# 8. JAHRESVERGLEICH (nur im Beide-Modus sichtbar)
# ============================================================
if compare_mode:
    st.header("8 · Jahresvergleich (Wachstum & Rückgang)")
    years_sorted = sorted(df["Jahr"].unique())
    ya, yb = years_sorted[0], years_sorted[-1]
    st.caption(f"Veränderung {yb} gegenüber {ya}")

    def _change_table(level: str, top_n: int = 15) -> pd.DataFrame:
        pivot = (
            df.groupby([level, "Jahr"])[metric].sum().unstack(fill_value=0)
        )
        for y in (ya, yb):
            if y not in pivot.columns:
                pivot[y] = 0
        pivot["Δ absolut"] = pivot[yb] - pivot[ya]
        pivot["Δ %"] = np.where(
            pivot[ya] > 0,
            (pivot[yb] - pivot[ya]) / pivot[ya] * 100,
            np.nan,
        )
        return pivot

    tab_g_ein, tab_g_ana = st.tabs(["Einsender", "Analysen/Leistungen"])

    for tab, level in [(tab_g_ein, "Einsender"), (tab_g_ana, "Analyse/Leistung")]:
        with tab:
            change = _change_table(level)
            top_abs = change.reindex(change["Δ absolut"].abs().sort_values(ascending=False).index).head(20)
            # Balkendiagramm: absolute Veränderung, diverging rot/grün
            fig_change = go.Figure()
            fig_change.add_trace(go.Bar(
                y=top_abs.index, x=top_abs["Δ absolut"], orientation="h",
                marker_color=["#2ca02c" if v >= 0 else "#d62728" for v in top_abs["Δ absolut"]],
                hovertemplate=("%{y}<br>"
                               f"{ya}: " + ("€\u202f%{customdata[0]:,.2f}" if _cost_mode else "%{customdata[0]:,.0f}") + "<br>"
                               f"{yb}: " + ("€\u202f%{customdata[1]:,.2f}" if _cost_mode else "%{customdata[1]:,.0f}") + "<br>"
                               "Δ: " + ("€\u202f%{x:,.2f}" if _cost_mode else "%{x:,.0f}") +
                               "<br>Δ %: %{customdata[2]:.1f}<extra></extra>"),
                customdata=np.stack([top_abs[ya].values, top_abs[yb].values, top_abs["Δ %"].fillna(0).values], axis=-1),
            ))
            fig_change.update_layout(
                title=f"Top 20 {level} nach absoluter Veränderung ({metric})",
                yaxis=dict(autorange="reversed"), height=max(400, len(top_abs) * 26),
                xaxis=dict(title=f"Δ {metric}", **_tick_fmt),
                margin=dict(l=10, r=10, t=60, b=40),
            )
            st.plotly_chart(fig_change, use_container_width=True)

            # Tabelle: sortierbar
            show = change[[ya, yb, "Δ absolut", "Δ %"]].copy()
            show = show.sort_values("Δ absolut", ascending=False)
            st.dataframe(
                show.style.format({
                    ya: ("{:,.2f} €" if _cost_mode else "{:,.0f}"),
                    yb: ("{:,.2f} €" if _cost_mode else "{:,.0f}"),
                    "Δ absolut": ("{:+,.2f} €" if _cost_mode else "{:+,.0f}"),
                    "Δ %": "{:+.1f} %",
                }),
                use_container_width=True, height=500,
            )

# ============================================================
# 9. DATEN-EXPLORER
# ============================================================
st.header(("9" if compare_mode else "8") + " · Daten-Explorer")

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
