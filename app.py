import streamlit as st
import pandas as pd
from datetime import datetime
from garminconnect import Garmin

# 1. CONFIGURAZIONE PAGINA PER CELLULARE
st.set_page_config(page_title="Luca Tassatori Coach", page_icon="👟", layout="centered")

st.markdown("""
    <style>
    /* Ottimizzazione spazi per schermi piccoli */
    .block-container { padding-top: 1rem; padding-bottom: 1rem; padding-left: 1rem; padding-right: 1rem; }
    .home-logo { text-align: center; font-size: 50px; margin-bottom: -10px; }
    .main-title { text-align: center; font-weight: 900; color: #1e293b; font-size: 2rem; margin-bottom: 20px;}
    
    /* Card Mobile */
    .card-planned { background-color: #f0f9ff; padding: 15px; border-radius: 12px; border-left: 5px solid #0ea5e9; margin-bottom: 15px;}
    .card-actual { background-color: #f0fdf4; padding: 15px; border-radius: 12px; border-left: 5px solid #22c55e; margin-bottom: 15px;}
    .card-title { font-size: 1rem; font-weight: bold; margin-bottom: 5px; color: #334155;}
    .workout-text { font-size: 1.1rem; margin:0; font-weight: 500;}
    </style>
""", unsafe_allow_html=True)

# Intestazione
st.markdown("<div class='home-logo'>👟</div>", unsafe_allow_html=True)
st.markdown("<h1 class='main-title'>Luca Tassatori Coach</h1>", unsafe_allow_html=True)

# Dizionari per gestire le date in italiano a prova di errore
mesi_abbreviati = {1:'gen', 2:'feb', 3:'mar', 4:'apr', 5:'mag', 6:'giu', 7:'lug', 8:'ago', 9:'set', 10:'ott', 11:'nov', 12:'dic'}
mesi_completi = {1:'gennaio', 2:'febbraio', 3:'marzo', 4:'aprile', 5:'maggio', 6:'giugno', 7:'luglio', 8:'agosto', 9:'settembre', 10:'ottobre', 11:'novembre', 12:'dicembre'}

oggi = datetime.today()
giorno_excel_oggi = f"{oggi.day}-{mesi_abbreviati[oggi.month]}" # Genera esattamente es. "24-lug"
mese_corrente_nome = mesi_completi[oggi.month] # Genera "luglio"

# Menu Laterale
st.sidebar.title("⚙️ Impostazioni")
uploaded_file = st.sidebar.file_uploader("1. Carica Excel Coach", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.subheader("2. Accesso Garmin Connect")
garmin_email = st.sidebar.text_input("Email Garmin")
garmin_pwd = st.sidebar.text_input("Password Garmin", type="password")

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        fogli_mesi = xls.sheet_names
        
        # Cerca qual è il foglio del mese corrente (es. cerca "luglio" dentro "luglio26")
        indice_mese_corrente = 0
        for i, nome_foglio in enumerate(fogli_mesi):
            if mese_corrente_nome in nome_foglio.lower():
                indice_mese_corrente = i
                break

        dati_completi = {}
        allenamento_oggi = "Riposo o nessun allenamento trovato per oggi."
        
        # Lettura di tutti i mesi
        for foglio in fogli_mesi:
            df = pd.read_excel(xls, sheet_name=foglio, header=None)
            row_dates = df.iloc[3].values if len(df) > 3 else []
            row_workouts = df.iloc[5].values if len(df) > 5 else []
            
            mese_data = []
            for i in range(1, len(row_dates)):
                d = str(row_dates[i]).strip()
                w = str(row_workouts[i]).strip()
                if d and d != "nan" and w and w != "nan":
                    mese_data.append({"Data": d, "Programma": w})
                    
                    # Controllo infallibile sulla data di oggi
                    if d.lower() == giorno_excel_oggi:
                        allenamento_oggi = w
            
            if mese_data:
                dati_completi[foglio] = pd.DataFrame(mese_data)

        # --- SEZIONE 1: OGGI ---
        st.subheader(f"⚡ Oggi ({oggi.strftime('%d/%m/%Y')})")
        
        st.markdown(f"""
        <div class="card-planned">
            <div class="card-title">🔵 Da Tabella Coach</div>
            <p class="workout-text">{allenamento_oggi}</p>
        </div>
        """, unsafe_allow_html=True)

        # --- INTEGRAZIONE GARMIN ---
        garmin_effettivo = "In attesa di credenziali Garmin..."
        if garmin_email and garmin_pwd:
            try:
                # Inizializza client Garmin
                client = Garmin(garmin_email, garmin_pwd)
                client.login()
                
                # Ottieni le attività di oggi
                oggi_iso = oggi.strftime('%Y-%m-%d')
                attivita = client.get_activities(0, 5) # Prende le ultime 5
                
                attivita_oggi = [a for a in attivita if a['startTimeLocal'].startswith(oggi_iso)]
                
                if attivita_oggi:
                    act = attivita_oggi[0] # Prende la prima di oggi
                    distanza = round(act.get('distance', 0) / 1000, 2)
                    durata_min = round(act.get('duration', 0) / 60, 1)
                    
                    # Calcolo passo (min/km)
                    velocita_ms = act.get('averageSpeed', 0)
                    if velocita_ms > 0:
                        passo_sec_km = 1000 / velocita_ms
                        minuti = int(passo_sec_km // 60)
                        secondi = int(passo_sec_km % 60)
                        passo_str = f"{minuti}'{secondi:02d}\""
                    else:
                        passo_str = "N/A"
                        
                    garmin_effettivo = f"**{act.get('activityName', 'Corsa')}**<br>📏 {distanza} km | ⏱️ {durata_min} min | ⚡ {passo_str}/km"
                else:
                    garmin_effettivo = "Nessuna attività registrata oggi sul Garmin."
            except Exception as e:
                garmin_effettivo = f"⚠️ Errore login Garmin. Controlla i dati."

        st.markdown(f"""
        <div class="card-actual">
            <div class="card-title">🟢 Svolto su Garmin</div>
            <p class="workout-text">{garmin_effettivo}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        # --- SEZIONE 2: STORICO OTTIMIZZATO PER CELLULARE ---
        st.subheader("📚 Storico Mensile")
        if dati_completi:
            # Menu a tendina che parte già dal mese corrente
            mese_scelto = st.selectbox("Seleziona il mese da visualizzare:", list(dati_completi.keys()), index=indice_mese_corrente)
            st.dataframe(dati_completi[mese_scelto], use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file: {e}")

else:
    st.info("👈 Apri il menu a tendina in alto a sinistra (>) e carica il file Excel per iniziare.")
