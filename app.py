import streamlit as st
import pandas as pd
from datetime import datetime
from garminconnect import Garmin

# 1. CONFIGURAZIONE PAGINA (Ottimizzata per evitare problemi di zoom su mobile)
st.set_page_config(page_title="Luca Tassatori Coach", page_icon="👟", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* Regolazioni margini e zoom per schermi smartphone */
    .block-container { padding: 1.5rem 1rem !important; max-width: 100%; overflow-x: hidden; }
    
    .home-logo { text-align: center; font-size: 45px; margin-bottom: -15px; }
    .main-title { text-align: center; font-weight: 900; color: #1e293b; font-size: 1.8rem; margin-bottom: 25px; line-height: 1.2;}
    
    /* Design Card con testi che non sbrodolano */
    .card-planned { background-color: #f0f9ff; padding: 15px; border-radius: 10px; border-left: 5px solid #0ea5e9; margin-bottom: 15px;}
    .card-actual { background-color: #f0fdf4; padding: 15px; border-radius: 10px; border-left: 5px solid #22c55e; margin-bottom: 15px;}
    .card-title { font-size: 0.95rem; font-weight: 700; margin-bottom: 8px; color: #475569; text-transform: uppercase; letter-spacing: 0.5px;}
    .workout-text { font-size: 1.05rem; margin:0; font-weight: 600; color: #0f172a;}
    
    /* Adatta le tabelle allo schermo mobile */
    [data-testid="stDataFrame"] { width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 2. INTESTAZIONE E CALENDARIO
st.markdown("<div class='home-logo'>👟</div>", unsafe_allow_html=True)
st.markdown("<h1 class='main-title'>Luca Tassatori Coach</h1>", unsafe_allow_html=True)

mesi_abbreviati = {1:'gen', 2:'feb', 3:'mar', 4:'apr', 5:'mag', 6:'giu', 7:'lug', 8:'ago', 9:'set', 10:'ott', 11:'nov', 12:'dic'}
oggi = datetime.today()
giorno_excel_oggi_testo = f"{oggi.day}-{mesi_abbreviati[oggi.month]}"

# 3. MENU LATERALE
st.sidebar.title("⚙️ Impostazioni")
uploaded_file = st.sidebar.file_uploader("1. Carica Excel Coach", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.subheader("2. Accesso Garmin")
garmin_email = st.sidebar.text_input("Email Garmin Connect")
garmin_pwd = st.sidebar.text_input("Password Garmin", type="password")

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # ORDINA I FOGLI: Tutti quelli che finiscono con "26" vanno per primi
        fogli_mesi = xls.sheet_names
        fogli_ordinati = sorted(fogli_mesi, key=lambda x: not str(x).strip().endswith('26'))

        dati_completi = {}
        allenamento_oggi = "Riposo o nessun allenamento previsto."
        
        # 4. LETTURA FILE CON CONTROLLO DATE AVANZATO
        for foglio in fogli_ordinati:
            df = pd.read_excel(xls, sheet_name=foglio, header=None)
            
            row_dates = df.iloc[3].values if len(df) > 3 else []
            row_workouts = df.iloc[5].values if len(df) > 5 else []
            
            mese_data = []
            for i in range(1, len(row_dates)):
                d_raw = row_dates[i]
                w = str(row_workouts[i]).strip()
                
                # Se la cella della data non è vuota e c'è un allenamento
                if pd.notna(d_raw) and w and w != "nan":
                    is_oggi = False
                    
                    # CASO A: Excel l'ha salvata come data vera e propria
                    if isinstance(d_raw, datetime) or type(d_raw).__name__ == 'Timestamp':
                        d_str = f"{d_raw.day}-{mesi_abbreviati.get(d_raw.month, '')}"
                        if d_raw.day == oggi.day and d_raw.month == oggi.month:
                            is_oggi = True
                    
                    # CASO B: Excel l'ha salvata come semplice testo
                    else:
                        d_str = str(d_raw).strip().lower()
                        if d_str == giorno_excel_oggi_testo:
                            is_oggi = True
                            
                    mese_data.append({"Data": d_str.capitalize(), "Programma": w})
                    
                    # Imposta l'allenamento di oggi se c'è corrispondenza
                    if is_oggi:
                        allenamento_oggi = w
            
            if mese_data:
                dati_completi[foglio] = pd.DataFrame(mese_data)

        # --- SEZIONE OGGI ---
        st.subheader(f"⚡ Oggi ({oggi.strftime('%d/%m/%Y')})")
        
        st.markdown(f"""
        <div class="card-planned">
            <div class="card-title">🔵 Da Tabella Coach</div>
            <p class="workout-text">{allenamento_oggi}</p>
        </div>
        """, unsafe_allow_html=True)

        # --- CONNESSIONE GARMIN ---
        garmin_effettivo = "Credenziali Garmin non inserite o sincronizzazione in corso..."
        if garmin_email and garmin_pwd:
            try:
                client = Garmin(garmin_email, garmin_pwd)
                client.login()
                
                oggi_iso = oggi.strftime('%Y-%m-%d')
                attivita = client.get_activities(0, 5)
                
                attivita_oggi = [a for a in attivita if a.get('startTimeLocal', '').startswith(oggi_iso)]
                
                if attivita_oggi:
                    act = attivita_oggi[0]
                    distanza = round(act.get('distance', 0) / 1000, 2)
                    durata_min = round(act.get('duration', 0) / 60, 1)
                    
                    velocita_ms = act.get('averageSpeed', 0)
                    if velocita_ms > 0:
                        passo_sec_km = 1000 / velocita_ms
                        passo_str = f"{int(passo_sec_km // 60)}'{int(passo_sec_km % 60):02d}\""
                    else:
                        passo_str = "N/A"
                        
                    garmin_effettivo = f"**{act.get('activityName', 'Corsa')}**<br>📏 {distanza} km &nbsp;|&nbsp; ⏱️ {durata_min} min &nbsp;|&nbsp; ⚡ {passo_str}/km"
                else:
                    garmin_effettivo = "Nessuna attività registrata oggi su Garmin."
            except Exception as e:
                garmin_effettivo = f"⚠️ Errore di accesso a Garmin. Verifica le credenziali."

        st.markdown(f"""
        <div class="card-actual">
            <div class="card-title">🟢 Svolto su Garmin</div>
            <p class="workout-text">{garmin_effettivo}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        # --- SEZIONE STORICO MESE (Tendina ordinata) ---
        st.subheader("📚 Storico Allenamenti")
        if dati_completi:
            # I fogli sono già ordinati (es. prima luglio26)
            mese_scelto = st.selectbox("Seleziona periodo:", list(dati_completi.keys()))
            st.dataframe(dati_completi[mese_scelto], use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"C'è stato un problema nella lettura dell'Excel. Assicurati che le date siano sulla riga giusta. Errore tecnico: {e}")

else:
    st.info("👈 Apri il menu laterale in alto a sinistra per caricare l'Excel e inserire i dati Garmin.")
