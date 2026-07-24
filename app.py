import streamlit as st
import pandas as pd
import altair as alt
import re
import calendar
from datetime import datetime, timedelta
from garminconnect import Garmin, GarminConnectAuthenticationError
import os

# 1. CONFIGURAZIONE PAGINA
st.set_page_config(page_title="Luca Tassarotti Coach", page_icon="👟", layout="centered", initial_sidebar_state="collapsed")

# ---------------------------------------------------------------------------
# PALETTE (a tema corsa/pista) — cambia solo qui per ridefinire i colori
# ---------------------------------------------------------------------------
COLORE_ASFALTO = "#1e293b"    # testo principale, sfondo scuro
COLORE_PISTA = "#ea580c"      # dati Garmin / svolto
COLORE_PISTA_CHIARO = "#fed7aa"
COLORE_CIELO = "#0ea5e9"      # piano del coach / pianificato
COLORE_TRAGUARDO = "#16a34a"  # SOLO per badge/traguardi raggiunti
COLORE_NEBBIA = "#f1f5f9"     # sfondo neutro / giorni senza attività

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;700&family=Inter:wght@400;500;600&display=swap');

    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    .block-container {{ padding: 1.5rem 1rem !important; max-width: 100%; overflow-x: hidden; }}
    .home-logo {{ text-align: center; font-size: 45px; margin-bottom: -15px; }}
    .main-title {{ text-align: center; font-family: 'Oswald', sans-serif; font-weight: 700; letter-spacing: 1px;
                   color: {COLORE_ASFALTO}; font-size: 1.9rem; margin-bottom: 5px; line-height: 1.2; text-transform: uppercase;}}
    .motivation-box {{ background-color: #f8fafc; border-left: 4px solid {COLORE_ASFALTO}; padding: 10px 15px; border-radius: 6px;
                       font-style: italic; color: #475569; font-size: 0.95rem; text-align: center; margin-bottom: 20px;}}
    .card-planned {{ background-color: #f0f9ff; padding: 15px; border-radius: 10px; border-left: 5px solid {COLORE_CIELO}; margin-bottom: 15px;}}
    .card-actual {{ background-color: #fff7ed; padding: 15px; border-radius: 10px; margin-bottom: 15px;}}
    .card-title {{ font-size: 0.95rem; font-weight: 700; margin-bottom: 8px; color: #475569; text-transform: uppercase;}}
    .workout-text {{ font-size: 1.05rem; margin:0; font-weight: 600; color: #0f172a;}}
    hr.garmin-divider {{ margin: 10px 0; border: 0; border-top: 1px solid #fed7aa; }}
    [data-testid="stDataFrame"] {{ width: 100%; }}

    .stat-number {{ font-family: 'Oswald', sans-serif; font-weight: 700; font-size: 1.9rem; color: {COLORE_ASFALTO}; line-height: 1;}}
    .stat-label {{ font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.5px; color: #64748b; margin-top: 2px;}}

    .badge-pill {{ display: inline-block; background-color: {COLORE_TRAGUARDO}; color: white; padding: 4px 12px;
                   border-radius: 999px; font-size: 0.8rem; font-weight: 600; margin: 3px 5px 3px 0;}}

    .progress-caption {{ font-size: 0.8rem; color: #78350f; margin-top: 8px; margin-bottom: 2px; }}
    .progress-track {{ background-color: {COLORE_NEBBIA}; border-radius: 999px; height: 10px; overflow: hidden; }}
    .progress-fill {{ height: 100%; border-radius: 999px; background-color: {COLORE_PISTA}; }}
    .progress-fill.completo {{ background-color: {COLORE_TRAGUARDO}; }}

    .heatmap-wrapper {{ max-width: 300px; margin: 8px auto 0 auto; }}
    .heatmap-weekdays {{ display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px; }}
    .heatmap-weekday {{ font-size: 0.6rem; text-align: center; color: #94a3b8; text-transform: uppercase; }}
    .heatmap-grid {{ display: grid; grid-template-columns: repeat(7, 1fr); gap: 3px; margin-top: 4px; }}
    .heatmap-cell {{ aspect-ratio: 1; border-radius: 5px; display: flex; flex-direction: column; align-items: center;
                     justify-content: center; line-height: 1.05; color: {COLORE_ASFALTO}; }}
    .heatmap-cell .hm-day {{ font-size: 0.6rem; font-weight: 600; }}
    .heatmap-cell .hm-km {{ font-size: 0.5rem; color: #7c2d12; }}
    .heatmap-cell.oggi {{ outline: 2px solid {COLORE_ASFALTO}; outline-offset: -2px; }}
    .heatmap-cell.vuota {{ visibility: hidden; }}
    </style>
""", unsafe_allow_html=True)

# 2. CREDENZIALI
GARMIN_EMAIL = "scocla@hotmail.it"
GARMIN_PWD = "Ciccio1994"

EXCEL_FILE_PATH = "storico_salvato.xlsx"

FRASI_MOTIVAZIONALI = [
    "“La fatica di oggi sarà il ritmo di gara di domani.”",
    "“Non correre solo con le gambe, corri con la testa e con il cuore.”",
    "“Il dolore è temporaneo, la gloria di aver finito è per sempre.”",
    "“Ogni passo avanti è un passo più vicino al tuo obiettivo.”",
    "“Se vuoi vincere qualcosa, corri 100 metri. Se vuoi goderti un'altra vita, corri la maratona.”",
    "“Il miracolo non è che ho terminato. Il miracolo è che ho avuto il coraggio di iniziare.”",
]

MESI_COMPLETI = {1: 'gennaio', 2: 'febbraio', 3: 'marzo', 4: 'aprile', 5: 'maggio', 6: 'giugno',
                 7: 'luglio', 8: 'agosto', 9: 'settembre', 10: 'ottobre', 11: 'novembre', 12: 'dicembre'}
MESI_IT = {1: 'gen', 2: 'feb', 3: 'mar', 4: 'apr', 5: 'mag', 6: 'giu', 7: 'lug', 8: 'ago',
           9: 'set', 10: 'ott', 11: 'nov', 12: 'dic'}
GIORNI_SETTIMANA_IT = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]

REGEX_KM = re.compile(r'(\d+(?:[.,]\d+)?)\s*km', re.IGNORECASE)


# ---------------------------------------------------------------------------
# GARMIN
# ---------------------------------------------------------------------------

@st.cache_resource(ttl=60 * 30)  # ricrea la sessione al massimo ogni 30 minuti
def get_garmin_client(email: str, password: str):
    client = Garmin(email, password)
    client.login()
    return client


@st.cache_data(ttl=60 * 10)
def sincronizza_garmin_periodo(_client, email_key: str, giorni: int = 35):
    """Recupera le attività degli ultimi N giorni (default 35, copre mese + streak + settimana).
    _client non viene hashato da streamlit (prefisso _)."""
    fine = datetime.today().date()
    inizio = fine - timedelta(days=giorni - 1)
    try:
        attivita = _client.get_activities_by_date(inizio.isoformat(), fine.isoformat())
    except AttributeError:
        # fallback per versioni di garminconnect senza questo metodo
        attivita = _client.get_activities(0, 60)
        attivita = [a for a in attivita if inizio.isoformat() <= a.get('startTimeLocal', '')[:10] <= fine.isoformat()]
    return attivita


@st.cache_data(ttl=60 * 30)
def ottieni_percorso_gps(_client, activity_id):
    """Prova a recuperare la polilinea GPS di un'attività. Ritorna None se non disponibile
    (dipende dalla versione della libreria garminconnect e dal tipo di attività)."""
    try:
        dettagli = _client.get_activity_details(activity_id)
        punti = dettagli.get("geoPolylineDTO", {}).get("polyline", [])
        coords = [(p["lat"], p["lon"]) for p in punti if "lat" in p and "lon" in p]
        return coords if coords else None
    except Exception:
        return None


def formatta_passo(velocita_ms: float) -> str:
    if not velocita_ms or velocita_ms <= 0:
        return "N/A"
    secondi_per_km = 1000 / velocita_ms
    minuti = int(secondi_per_km // 60)
    secondi = int(secondi_per_km % 60)
    return f"{minuti}'{secondi:02d}\""


def raggruppa_per_giorno(attivita: list) -> dict:
    """Ritorna {data_iso: {'km': totale, 'attivita': [...]}}"""
    per_giorno = {}
    for act in attivita:
        data_iso = act.get('startTimeLocal', '')[:10]
        if not data_iso:
            continue
        km = act.get('distance', 0) / 1000
        per_giorno.setdefault(data_iso, {"km": 0.0, "attivita": []})
        per_giorno[data_iso]["km"] += km
        per_giorno[data_iso]["attivita"].append(act)
    return per_giorno


def calcola_streak(attivita_per_giorno: dict, oggi: datetime) -> int:
    """Conta i giorni consecutivi con almeno un'attività, partendo da oggi a ritroso."""
    streak = 0
    giorno = oggi.date()
    while True:
        info = attivita_per_giorno.get(giorno.isoformat())
        if info and info["km"] > 0:
            streak += 1
            giorno -= timedelta(days=1)
        else:
            break
    return streak


def statistiche_mese(attivita: list, oggi: datetime) -> dict:
    """Km totali, record distanza singola e miglior passo nel mese corrente."""
    km_totali, record_distanza, record_velocita = 0.0, 0.0, 0.0
    for act in attivita:
        data_iso = act.get('startTimeLocal', '')[:10]
        if not data_iso or not data_iso.startswith(oggi.strftime("%Y-%m")):
            continue
        km = act.get('distance', 0) / 1000
        km_totali += km
        record_distanza = max(record_distanza, km)
        record_velocita = max(record_velocita, act.get('averageSpeed', 0) or 0)
    return {"km_totali": km_totali, "record_distanza": record_distanza, "record_passo": formatta_passo(record_velocita)}


# ---------------------------------------------------------------------------
# EXCEL: PARSING PIU' ROBUSTO
# ---------------------------------------------------------------------------

@st.cache_data(ttl=60 * 5)
def carica_fogli_excel(path: str, mtime: float):
    xls = pd.ExcelFile(path)
    return {foglio: pd.read_excel(xls, sheet_name=foglio, header=None) for foglio in xls.sheet_names}


def trova_riga_date(df: pd.DataFrame, mese_oggi: int) -> int:
    """Sceglie la riga con più celle che sembrano date, invece che la prima trovata."""
    migliore_idx, miglior_punteggio = -1, 0
    for r_idx in range(len(df)):
        punteggio = 0
        for val in df.iloc[r_idx].values:
            if isinstance(val, datetime) or type(val).__name__ == "Timestamp":
                punteggio += 1
                continue
            val_str = str(val).lower()
            if (f"-{MESI_IT[mese_oggi]}" in val_str or f"/{mese_oggi}/" in val_str
                    or "lug" in val_str or "jul" in val_str):
                punteggio += 1
        if punteggio > miglior_punteggio:
            miglior_punteggio, migliore_idx = punteggio, r_idx
    if migliore_idx == -1 and len(df) > 3:
        return 3
    return migliore_idx


def estrai_workout_del_mese(df: pd.DataFrame, mese_oggi: int, giorno_oggi: int):
    """Ritorna (righe_mese, allenamento_oggi, celle_non_riconosciute)."""
    mese_data = []
    allenamento_oggi = None
    non_riconosciute = 0

    riga_date_idx = trova_riga_date(df, mese_oggi)
    if riga_date_idx == -1:
        return mese_data, allenamento_oggi, non_riconosciute

    riga_workout_idx = riga_date_idx + 2 if (riga_date_idx + 2) < len(df) else riga_date_idx + 1
    if riga_workout_idx >= len(df):
        return mese_data, allenamento_oggi, non_riconosciute

    row_dates = df.iloc[riga_date_idx].values
    row_workouts = df.iloc[riga_workout_idx].values

    for col_idx in range(len(row_dates)):
        d_raw = row_dates[col_idx]
        w = str(row_workouts[col_idx]).strip() if col_idx < len(row_workouts) else ""

        if pd.isna(d_raw):
            continue

        d_str, is_oggi, riconosciuta = "", False, False

        if isinstance(d_raw, datetime) or type(d_raw).__name__ == "Timestamp":
            d_str = f"{d_raw.day}-{MESI_IT.get(d_raw.month, '')}"
            is_oggi = (d_raw.day == giorno_oggi and d_raw.month == mese_oggi)
            riconosciuta = True
        else:
            d_str = str(d_raw).strip()
            d_lower = d_str.lower()
            ha_mese = ("lug" in d_lower or "jul" in d_lower or str(mese_oggi) in d_lower)
            ha_giorno = any(ch.isdigit() for ch in d_lower)
            riconosciuta = ha_mese and ha_giorno
            is_oggi = (
                (f"{giorno_oggi}-" in d_lower or f"{giorno_oggi} " in d_lower or d_lower.startswith(str(giorno_oggi)))
                and ha_mese
            )

        if not riconosciuta:
            non_riconosciute += 1

        if d_str and d_str != "nan":
            mese_data.append({"Data": d_str.capitalize(), "Programma": w if w != "nan" else "Riposo"})

        if is_oggi and w and w not in ("nan", ""):
            allenamento_oggi = w

    mese_data.reverse()  # più recenti in alto, più vecchi in basso
    return mese_data, allenamento_oggi, non_riconosciute


def estrai_km_pianificati(testo: str):
    if not testo:
        return None
    match = REGEX_KM.search(testo)
    if match:
        return float(match.group(1).replace(",", "."))
    return None


# ---------------------------------------------------------------------------
# HEATMAP MENSILE (stile "contribution calendar")
# ---------------------------------------------------------------------------

def colore_heatmap(km: float) -> str:
    if km <= 0:
        return COLORE_NEBBIA
    if km < 5:
        return COLORE_PISTA_CHIARO
    if km < 10:
        return "#fb923c"
    return COLORE_PISTA


def costruisci_heatmap_html(oggi: datetime, attivita_per_giorno: dict) -> str:
    primo_weekday, giorni_nel_mese = calendar.monthrange(oggi.year, oggi.month)  # weekday: lunedì=0
    celle = ['<div class="heatmap-cell vuota"></div>' for _ in range(primo_weekday)]

    for giorno_num in range(1, giorni_nel_mese + 1):
        data = datetime(oggi.year, oggi.month, giorno_num)
        data_iso = data.strftime("%Y-%m-%d")
        km = attivita_per_giorno.get(data_iso, {}).get("km", 0.0)
        e_futuro = data.date() > oggi.date()
        colore = "white" if e_futuro else colore_heatmap(km)
        bordo = "border: 1px dashed #e2e8f0;" if e_futuro else ""
        classe_oggi = " oggi" if data.date() == oggi.date() else ""
        titolo = f"{giorno_num} {MESI_IT[oggi.month]}: {km:.1f} km" if not e_futuro else ""
        km_html = f'<span class="hm-km">{km:.0f}k</span>' if (not e_futuro and km > 0) else ""
        celle.append(
            f'<div class="heatmap-cell{classe_oggi}" style="background-color:{colore};{bordo}" title="{titolo}">'
            f'<span class="hm-day">{giorno_num}</span>{km_html}</div>'
        )

    intestazioni = "".join(f'<div class="heatmap-weekday">{g}</div>' for g in GIORNI_SETTIMANA_IT)
    return (
        f'<div class="heatmap-wrapper">'
        f'<div class="heatmap-weekdays">{intestazioni}</div>'
        f'<div class="heatmap-grid">{"".join(celle)}</div>'
        f'</div>'
    )


# ---------------------------------------------------------------------------
# STATO / DATE
# ---------------------------------------------------------------------------

oggi = datetime.today()
giorno_oggi = oggi.day
mese_oggi = oggi.month
oggi_iso = oggi.strftime("%Y-%m-%d")
foglio_default_target = f"{MESI_COMPLETI[mese_oggi]}{str(oggi.year)[-2:]}"

st.markdown("<div class='home-logo'>👟</div>", unsafe_allow_html=True)
st.markdown("<h1 class='main-title'>Luca Tassarotti Coach</h1>", unsafe_allow_html=True)

frase_del_giorno = FRASI_MOTIVAZIONALI[oggi.timetuple().tm_yday % len(FRASI_MOTIVAZIONALI)]
st.markdown(f"<div class='motivation-box'>{frase_del_giorno}</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------

st.sidebar.title("⚙️ Impostazioni")
uploaded_file = st.sidebar.file_uploader("Aggiorna File Excel", type=["xlsx", "xls"])
if uploaded_file is not None:
    with open(EXCEL_FILE_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.cache_data.clear()
    st.sidebar.success("File Excel aggiornato e salvato con successo!")

file_da_leggere = EXCEL_FILE_PATH if os.path.exists(EXCEL_FILE_PATH) else None

# ---------------------------------------------------------------------------
# GARMIN: sincronizzazione (35 giorni: copre mese corrente, streak e settimana)
# ---------------------------------------------------------------------------

attivita_periodo = []
attivita_per_giorno = {}
garmin_client = None
garmin_errore = None

if GARMIN_EMAIL and GARMIN_PWD:
    try:
        garmin_client = get_garmin_client(GARMIN_EMAIL, GARMIN_PWD)
        attivita_periodo = sincronizza_garmin_periodo(garmin_client, GARMIN_EMAIL)
        attivita_per_giorno = raggruppa_per_giorno(attivita_periodo)
    except GarminConnectAuthenticationError:
        garmin_errore = "auth"
    except Exception:
        garmin_errore = "connessione"

# ---------------------------------------------------------------------------
# ELABORAZIONE DATI EXCEL
# ---------------------------------------------------------------------------

if file_da_leggere:
    try:
        mtime = os.path.getmtime(file_da_leggere)
        fogli_raw = carica_fogli_excel(file_da_leggere, mtime)
        fogli_mesi = list(fogli_raw.keys())

        indice_default = next(
            (i for i, nome in enumerate(fogli_mesi) if foglio_default_target in nome.lower()), 0
        )

        dati_completi = {}
        allenamento_oggi = "Riposo o nessun allenamento trovato."
        totale_non_riconosciute = 0

        for foglio, df in fogli_raw.items():
            mese_data, trovato_oggi, non_ric = estrai_workout_del_mese(df, mese_oggi, giorno_oggi)
            if mese_data:
                dati_completi[foglio] = pd.DataFrame(mese_data)
            if trovato_oggi:
                allenamento_oggi = trovato_oggi
            totale_non_riconosciute += non_ric

        # --- ALLENAMENTO PROGRAMMATO ---
        st.subheader(f"⚡ Oggi ({oggi.strftime('%d/%m/%Y')})")
        st.markdown(f"""
        <div class="card-planned">
            <div class="card-title">🔵 Da Tabella Coach</div>
            <p class="workout-text">{allenamento_oggi}</p>
        </div>
        """, unsafe_allow_html=True)

        # --- ALLENAMENTO SVOLTO (solo Garmin) ---
        st.markdown("<div class='card-actual'>", unsafe_allow_html=True)
        st.markdown("<div class='card-title'>🟠 Allenamento Svolto</div>", unsafe_allow_html=True)

        km_svolto_oggi = 0.0

        if garmin_errore == "auth":
            st.error("❌ Credenziali Garmin non valide. Controlla email e password in secrets.toml.")
        elif garmin_errore == "connessione":
            st.markdown(
                "<p class='workout-text'>⚠️ Garmin temporaneamente non raggiungibile.</p>",
                unsafe_allow_html=True,
            )
        else:
            dati_oggi = attivita_per_giorno.get(oggi_iso)
            if dati_oggi and dati_oggi["attivita"]:
                km_svolto_oggi = dati_oggi["km"]
                for i, act in enumerate(dati_oggi["attivita"]):
                    distanza = round(act.get("distance", 0) / 1000, 2)
                    durata_min = round(act.get("duration", 0) / 60, 1)
                    passo_str = formatta_passo(act.get("averageSpeed", 0))
                    if i > 0:
                        st.markdown("<hr class='garmin-divider'>", unsafe_allow_html=True)
                    st.markdown(
                        f"<p class='workout-text'><strong>{act.get('activityName', 'Attività')}</strong><br>"
                        f"📏 {distanza} km &nbsp;|&nbsp; ⏱️ {durata_min} min &nbsp;|&nbsp; ⚡ {passo_str}/km</p>",
                        unsafe_allow_html=True,
                    )

                    # Mini mappa del percorso (se Garmin espone la polilinea GPS)
                    percorso = ottieni_percorso_gps(garmin_client, act.get("activityId"))
                    if percorso:
                        df_percorso = pd.DataFrame(percorso, columns=["lat", "lon"])
                        mappa = (
                            alt.Chart(df_percorso)
                            .mark_line(color=COLORE_PISTA, strokeWidth=3)
                            .encode(x=alt.X("lon:Q", axis=None, scale=alt.Scale(zero=False)),
                                    y=alt.Y("lat:Q", axis=None, scale=alt.Scale(zero=False)))
                            .properties(height=140)
                            .configure_view(strokeWidth=0)
                        )
                        st.altair_chart(mappa, use_container_width=True)
            else:
                st.markdown(
                    "<p class='workout-text'>Nessuna attività registrata oggi su Garmin.</p>",
                    unsafe_allow_html=True,
                )

            with st.expander("🗓️ Attività dell'ultima settimana"):
                giorni_settimana_iso = {(oggi - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(7)}
                trovate = False
                for data_iso in sorted(giorni_settimana_iso, reverse=True):
                    info = attivita_per_giorno.get(data_iso)
                    if info:
                        trovate = True
                        st.markdown(f"**{data_iso}** — {info['km']:.1f} km totali ({len(info['attivita'])} attività)")
                if not trovate:
                    st.caption("Nessuna attività trovata nell'ultima settimana.")

        # Barra di progresso: km svolti oggi rispetto al piano (senza etichette tipo "target")
        km_piano_oggi = estrai_km_pianificati(allenamento_oggi)
        if km_piano_oggi and km_piano_oggi > 0:
            percentuale = min(100, round((km_svolto_oggi / km_piano_oggi) * 100))
            classe_fill = "completo" if percentuale >= 100 else ""
            st.markdown(f"<p class='progress-caption'>{km_svolto_oggi:.1f} / {km_piano_oggi:.1f} km di oggi</p>", unsafe_allow_html=True)
            st.markdown(
                f"<div class='progress-track'><div class='progress-fill {classe_fill}' style='width:{percentuale}%;'></div></div>",
                unsafe_allow_html=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

        if totale_non_riconosciute > 0:
            st.caption(f"ℹ️ {totale_non_riconosciute} celle data non riconosciute nell'Excel (ignorate nel parsing).")

        st.markdown("---")

        # --- STATISTICHE DEL MESE + STREAK + BADGE ---
        stats = statistiche_mese(attivita_periodo, oggi)
        streak = calcola_streak(attivita_per_giorno, oggi)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"<div class='stat-number'>{stats['km_totali']:.0f}</div><div class='stat-label'>Km questo mese</div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='stat-number'>{streak}</div><div class='stat-label'>Giorni di fila</div>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<div class='stat-number'>{stats['record_distanza']:.1f}</div><div class='stat-label'>Km più lunghi</div>", unsafe_allow_html=True)

        badge = []
        if stats["km_totali"] >= 100:
            badge.append("🏅 100 km questo mese")
        if streak >= 5:
            badge.append(f"🔥 {streak} allenamenti di fila")
        if stats["record_distanza"] >= 15:
            badge.append(f"🏆 Lunga distanza: {stats['record_distanza']:.1f} km")

        if badge:
            st.markdown(
                "".join(f"<span class='badge-pill'>{b}</span>" for b in badge),
                unsafe_allow_html=True,
            )

        st.markdown("---")

        # --- HEATMAP MENSILE ---
        st.subheader("🔥 Calendario del mese")
        st.markdown(costruisci_heatmap_html(oggi, attivita_per_giorno), unsafe_allow_html=True)

        st.markdown("---")

        # --- GRAFICO SETTIMANALE: ANDAMENTO DEL PASSO ---
        st.subheader("⚡ Andamento del passo (ultimi 7 giorni)")
        giorni_settimana = [(oggi - timedelta(days=i)) for i in range(6, -1, -1)]

        righe_passo = []
        for giorno in giorni_settimana:
            g_iso = giorno.strftime("%Y-%m-%d")
            info = attivita_per_giorno.get(g_iso)
            passo_min_km = None
            if info and info["attivita"]:
                distanza_tot = sum(a.get("distance", 0) for a in info["attivita"])
                durata_tot = sum(a.get("duration", 0) for a in info["attivita"])
                if distanza_tot > 0 and durata_tot > 0:
                    velocita_media = distanza_tot / durata_tot  # m/s
                    passo_min_km = (1000 / velocita_media) / 60  # minuti per km

            righe_passo.append({"Giorno": giorno.strftime("%d/%m"), "PassoMinKm": passo_min_km})

        df_passo = pd.DataFrame(righe_passo)
        if df_passo["PassoMinKm"].notna().any():
            grafico_passo = (
                alt.Chart(df_passo.dropna(subset=["PassoMinKm"]))
                .mark_line(point=True, color=COLORE_PISTA, strokeWidth=3)
                .encode(
                    x=alt.X("Giorno:N", sort=list(df_passo["Giorno"]), title=None),
                    y=alt.Y("PassoMinKm:Q", title="min/km", scale=alt.Scale(zero=False)),
                    tooltip=[alt.Tooltip("Giorno:N"), alt.Tooltip("PassoMinKm:Q", title="min/km", format=".2f")],
                )
                .properties(height=220)
            )
            st.altair_chart(grafico_passo, use_container_width=True)
        else:
            st.caption("Non ci sono ancora attività Garmin questa settimana per mostrare il passo.")

        st.markdown("---")

        # --- STORICO ---
        st.subheader("📚 Storico Allenamenti")
        if dati_completi:
            mese_scelto = st.selectbox("Seleziona periodo:", list(dati_completi.keys()), index=indice_default)
            tabella_storico = dati_completi[mese_scelto]
            etichetta_oggi = f"{giorno_oggi}-{MESI_IT[mese_oggi]}".capitalize()

            def evidenzia_oggi(row):
                if row["Data"] == etichetta_oggi:
                    return ["background-color: #fed7aa; font-weight: 700;"] * len(row)
                return [""] * len(row)

            tabella_stilizzata = tabella_storico.style.apply(evidenzia_oggi, axis=1)
            st.dataframe(tabella_stilizzata, use_container_width=True, hide_index=True)
        else:
            st.info("Nessun dato leggibile trovato nell'Excel per i mesi caricati.")

    except Exception as e:
        st.error(f"Errore tecnico durante la lettura dell'Excel: {e}")

else:
    st.info("👈 Apri il menu laterale in alto a sinistra per caricare l'Excel per la prima volta.")
