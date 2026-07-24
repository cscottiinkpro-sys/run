import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
from garminconnect import Garmin, GarminConnectAuthenticationError
import os

# 1. CONFIGURAZIONE PAGINA
st.set_page_config(page_title="Luca Tassarotti Coach", page_icon="👟", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .block-container { padding: 1.5rem 1rem !important; max-width: 100%; overflow-x: hidden; }
    .home-logo { text-align: center; font-size: 45px; margin-bottom: -15px; }
    .main-title { text-align: center; font-weight: 900; color: #1e293b; font-size: 1.8rem; margin-bottom: 5px; line-height: 1.2;}
    .motivation-box { background-color: #f8fafc; border-left: 4px solid #64748b; padding: 10px 15px; border-radius: 6px; font-style: italic; color: #475569; font-size: 0.95rem; text-align: center; margin-bottom: 20px;}
    .card-planned { background-color: #f0f9ff; padding: 15px; border-radius: 10px; border-left: 5px solid #0ea5e9; margin-bottom: 15px;}
    .card-actual { background-color: #f0fdf4; padding: 15px; border-radius: 10px; border-left: 5px solid #22c55e; margin-bottom: 15px;}
    .card-title { font-size: 0.95rem; font-weight: 700; margin-bottom: 8px; color: #475569; text-transform: uppercase;}
    .workout-text { font-size: 1.05rem; margin:0; font-weight: 600; color: #0f172a;}
    .badge-ok { color: #16a34a; font-weight: 700; }
    .badge-warn { color: #ca8a04; font-weight: 700; }
    .badge-bad { color: #dc2626; font-weight: 700; }
    hr.garmin-divider { margin: 10px 0; border: 0; border-top: 1px solid #bbf7d0; }
    [data-testid="stDataFrame"] { width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 2. CREDENZIALI
GARMIN_EMAIL = "scocla@hotmail.it"
GARMIN_PWD = "Ciccio1994"

EXCEL_FILE_PATH = "storico_salvato.xlsx"
NOTE_LOG_PATH = "note_allenamenti_log.csv"

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
def sincronizza_garmin_settimana(_client, email_key: str, giorni: int = 7):
    """Recupera le attività dell'ultima settimana. _client non viene hashato da streamlit (prefisso _)."""
    fine = datetime.today().date()
    inizio = fine - timedelta(days=giorni - 1)
    try:
        attivita = _client.get_activities_by_date(inizio.isoformat(), fine.isoformat())
    except AttributeError:
        # fallback per versioni di garminconnect senza questo metodo
        attivita = _client.get_activities(0, 25)
        attivita = [a for a in attivita if inizio.isoformat() <= a.get('startTimeLocal', '')[:10] <= fine.isoformat()]
    return attivita


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


# ---------------------------------------------------------------------------
# NOTE MANUALI (STORICO)
# ---------------------------------------------------------------------------

def carica_note_log() -> pd.DataFrame:
    if os.path.exists(NOTE_LOG_PATH):
        try:
            df = pd.read_csv(NOTE_LOG_PATH, dtype=str).fillna("")
            if "Data" in df.columns and "Nota" in df.columns:
                return df
        except Exception:
            pass
    return pd.DataFrame(columns=["Data", "Nota"])


def salva_nota_log(data_iso: str, testo: str) -> pd.DataFrame:
    df = carica_note_log()
    df = df[df["Data"] != data_iso]
    nuova = pd.DataFrame([{"Data": data_iso, "Nota": testo}])
    df = pd.concat([df, nuova], ignore_index=True)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df.sort_values("Data", ascending=False)
    df["Data"] = df["Data"].dt.strftime("%Y-%m-%d")
    df.to_csv(NOTE_LOG_PATH, index=False)
    return df


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


# ---------------------------------------------------------------------------
# CONFRONTO PIANIFICATO VS SVOLTO
# ---------------------------------------------------------------------------

def estrai_km_pianificati(testo: str):
    if not testo:
        return None
    match = REGEX_KM.search(testo)
    if match:
        return float(match.group(1).replace(",", "."))
    return None


def valuta_aderenza(km_piano, km_svolto):
    """Ritorna (classe_css, emoji, messaggio)."""
    if km_piano is None and km_svolto is None:
        return "", "", None
    if km_piano is None:
        return "badge-warn", "🟡", f"Corsi {km_svolto:.1f} km (nessun target riconosciuto nel piano)"
    if km_svolto is None or km_svolto == 0:
        return "badge-bad", "🔴", f"Piano: {km_piano:.1f} km — nessuna attività registrata"

    diff_pct = abs(km_svolto - km_piano) / km_piano if km_piano else 1
    if diff_pct <= 0.10:
        return "badge-ok", "🟢", f"In linea col piano ({km_svolto:.1f} km vs {km_piano:.1f} km)"
    elif diff_pct <= 0.25:
        return "badge-warn", "🟡", f"Leggero scostamento ({km_svolto:.1f} km vs {km_piano:.1f} km)"
    else:
        return "badge-bad", "🔴", f"Scostamento marcato ({km_svolto:.1f} km vs {km_piano:.1f} km)"


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
# GARMIN: sincronizzazione settimanale (usata sia per oggi che per il grafico)
# ---------------------------------------------------------------------------

attivita_per_giorno = {}
garmin_client = None
garmin_errore = None

if GARMIN_EMAIL and GARMIN_PWD:
    try:
        garmin_client = get_garmin_client(GARMIN_EMAIL, GARMIN_PWD)
        attivita_settimana = sincronizza_garmin_settimana(garmin_client, GARMIN_EMAIL)
        attivita_per_giorno = raggruppa_per_giorno(attivita_settimana)
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

        # --- ALLENAMENTO SVOLTO ---
        st.markdown("<div class='card-actual'>", unsafe_allow_html=True)
        st.markdown("<div class='card-title'>🟢 Allenamento Svolto</div>", unsafe_allow_html=True)

        modalita = st.radio(
            "Scegli modalità:",
            ["Scrivi a mano", "Sincronizza da Garmin"],
            horizontal=True,
            label_visibility="collapsed",
        )

        if modalita == "Scrivi a mano":
            note_df = carica_note_log()
            nota_esistente = ""
            match_oggi = note_df[note_df["Data"] == oggi_iso]
            if not match_oggi.empty:
                nota_esistente = match_oggi.iloc[0]["Nota"]

            user_note = st.text_area(
                "Inserisci km, ritmi o sensazioni di oggi:", value=nota_esistente, height=80, key="input_manuale"
            )
            if st.button("Salva Nota"):
                salva_nota_log(oggi_iso, user_note)
                st.success("Salvato!")

            with st.expander("📜 Storico note"):
                storico = carica_note_log()
                if not storico.empty:
                    st.dataframe(storico, use_container_width=True, hide_index=True)
                else:
                    st.caption("Nessuna nota salvata finora.")

        else:
            if garmin_errore == "auth":
                st.error("❌ Credenziali Garmin non valide. Controlla email e password in secrets.toml.")
            elif garmin_errore == "connessione":
                st.markdown(
                    "<p class='workout-text'>⚠️ Garmin temporaneamente non raggiungibile. "
                    "Usa la modalità manuale qui sopra se hai fretta!</p>",
                    unsafe_allow_html=True,
                )
            else:
                dati_oggi = attivita_per_giorno.get(oggi_iso)
                if dati_oggi and dati_oggi["attivita"]:
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
                else:
                    st.markdown(
                        "<p class='workout-text'>Nessuna attività registrata oggi su Garmin.</p>",
                        unsafe_allow_html=True,
                    )

                with st.expander("🗓️ Attività dell'ultima settimana"):
                    if attivita_per_giorno:
                        for data_iso in sorted(attivita_per_giorno.keys(), reverse=True):
                            info = attivita_per_giorno[data_iso]
                            st.markdown(f"**{data_iso}** — {info['km']:.1f} km totali ({len(info['attivita'])} attività)")
                    else:
                        st.caption("Nessuna attività trovata nell'ultima settimana.")

        st.markdown("</div>", unsafe_allow_html=True)

        if totale_non_riconosciute > 0:
            st.caption(f"ℹ️ {totale_non_riconosciute} celle data non riconosciute nell'Excel (ignorate nel parsing).")

        st.markdown("---")

        # --- GRAFICO SETTIMANALE: PIANIFICATO VS SVOLTO ---
        st.subheader("📊 Andamento settimanale")
        foglio_corrente = dati_completi.get(fogli_mesi[indice_default]) if fogli_mesi else None
        giorni_settimana = [(oggi - timedelta(days=i)) for i in range(6, -1, -1)]

        righe_grafico = []
        for giorno in giorni_settimana:
            g_iso = giorno.strftime("%Y-%m-%d")
            km_svolto = attivita_per_giorno.get(g_iso, {}).get("km", 0.0)

            km_piano = 0.0
            if foglio_corrente is not None:
                etichetta = f"{giorno.day}-{MESI_IT[giorno.month]}".capitalize()
                riga = foglio_corrente[foglio_corrente["Data"] == etichetta]
                if not riga.empty:
                    km_piano = estrai_km_pianificati(riga.iloc[0]["Programma"]) or 0.0

            righe_grafico.append({"Giorno": giorno.strftime("%d/%m"), "Pianificato": km_piano, "Svolto": km_svolto})

        df_grafico = pd.DataFrame(righe_grafico).set_index("Giorno")
        if df_grafico[["Pianificato", "Svolto"]].sum().sum() > 0:
            st.bar_chart(df_grafico)
        else:
            st.caption("Non ci sono ancora abbastanza dati (piano o Garmin) per mostrare il grafico.")

        st.markdown("---")

        # --- STORICO ---
        st.subheader("📚 Storico Allenamenti")
        if dati_completi:
            mese_scelto = st.selectbox("Seleziona periodo:", list(dati_completi.keys()), index=indice_default)
            st.dataframe(dati_completi[mese_scelto], use_container_width=True, hide_index=True)
        else:
            st.info("Nessun dato leggibile trovato nell'Excel per i mesi caricati.")

    except Exception as e:
        st.error(f"Errore tecnico durante la lettura dell'Excel: {e}")

else:
    st.info("👈 Apri il menu laterale in alto a sinistra per caricare l'Excel per la prima volta.")
