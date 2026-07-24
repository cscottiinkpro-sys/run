import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# 1. CONFIGURAZIONE PAGINA E DESIGN PROFESSIONALE
st.set_page_config(page_title="Luca Tassatori Coach", page_icon="👟", layout="wide")

# CSS Personalizzato per Menu Laterale, Logo in Home e Card
st.markdown("""
    <style>
    /* Logo e Titolo in Home */
    .home-logo { text-align: center; font-size: 70px; margin-bottom: -15px; }
    .main-title { text-align: center; font-weight: 900; color: #1e293b; font-size: 2.5rem; }
    
    /* Menu Laterale ad alto contrasto */
    [data-testid="stSidebar"] { background-color: #0f172a; }
    [data-testid="stSidebar"] .css-17lntkn, [data-testid="stSidebar"] p, [data-testid="stSidebar"] div, [data-testid="stSidebar"] label { color: #f8fafc !important; }
    hr { margin-top: 1rem; margin-bottom: 1rem; border-color: #334155; }
    
    /* Stile delle Card di confronto */
    .card-planned { background-color: #f0f9ff; padding: 20px; border-radius: 12px; border-left: 6px solid #0ea5e9; height: 100%; box-shadow: 0 4px 6px rgba(0,0,0,0.05);}
    .card-actual { background-color: #fff1f2; padding: 20px; border-radius: 12px; border-left: 6px solid #f43f5e; height: 100%; box-shadow: 0 4px 6px rgba(0,0,0,0.05);}
    .card-title { font-size: 1.2rem; font-weight: bold; margin-bottom: 10px; color: #334155;}
    </style>
""", unsafe_allow_html=True)

# 2. INTESTAZIONE
st.markdown("<div class='home-logo'>👟</div>", unsafe_allow_html=True)
st.markdown("<h1 class='main-title'>Luca Tassatori Coach</h1>", unsafe_allow_html=True)
st.markdown("---")

# 3. MENU LATERALE STRUTTURATO
st.sidebar.title("⚙️ Impostazioni")

st.sidebar.subheader("1. Importa Storico Database")
uploaded_file = st.sidebar.file_uploader("Carica il file Excel (tutti i mesi)", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.subheader("2. Connessione Strava")
strava_token = st.sidebar.text_input("Strava Access Token", type="password")

# 4. LOGICA DI LETTURA E VISUALIZZAZIONE
oggi_str = datetime.today().strftime('%d-%b').lower() # Es. 24-lug

if uploaded_file is not None:
    try:
        # Legge tutti i fogli del file Excel
        xls = pd.ExcelFile(uploaded_file)
        fogli_mesi = xls.sheet_names
        
        dati_completi = {}
        allenamento_oggi = "Nessun allenamento programmato o data non trovata."
        
        # Estrazione dati per ogni mese (adattato alla struttura orizzontale)
        for foglio in fogli_mesi:
            df = pd.read_excel(xls, sheet_name=foglio, header=None)
            
            # Cerca le righe che contengono le date e gli allenamenti 
            # In base allo screenshot, supponiamo riga indice 3 o 4 per le date, e +2 per gli allenamenti
            # Qui usiamo un metodo flessibile che cerca la riga con più date e la successiva
            
            row_dates = df.iloc[3].values if len(df) > 3 else []
            row_workouts = df.iloc[5].values if len(df) > 5 else []
            
            mese_data = []
            for i in range(1, len(row_dates)):
                d = str(row_dates[i]).strip()
                w = str(row_workouts[i]).strip()
                if d and d != "nan" and w and w != "nan":
                    mese_data.append({"Data": d, "Programmato": w})
                    # Identifica l'allenamento di oggi
                    if d.lower() == oggi_str or d == datetime.today().strftime('%d-%m'):
                        allenamento_oggi = w
            
            if mese_data:
                dati_completi[foglio] = pd.DataFrame(mese_data)

        # --- SEZIONE: ALLENAMENTO DELLA GIORNATA ---
        st.subheader(f"⚡ Allenamento di Oggi ({datetime.today().strftime('%d/%m/%Y')})")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"""
            <div class="card-planned">
                <div class="card-title">🔵 Da Tabella Coach</div>
                <p style="font-size: 1.1rem; margin:0;">{allenamento_oggi}</p>
            </div>
            """, unsafe_allow_html=True)
            
        with col2:
            # Recupero dati da Strava per l'affiancamento
            strava_effettivo = "In attesa di sincronizzazione..."
            if strava_token:
                headers = {'Authorization': f'Bearer {strava_token}'}
                res = requests.get('https://www.strava.com/api/v3/athlete/activities?per_page=3', headers=headers)
                if res.status_code == 200:
                    acts = res.json()
                    # Cerca l'attività di oggi
                    oggi_iso = datetime.today().strftime('%Y-%m-%d')
                    attivita_oggi = [a for a in acts if a['start_date_local'].startswith(oggi_iso)]
                    
                    if attivita_oggi:
                        act = attivita_oggi[0]
                        distanza = round(act.get('distance', 0) / 1000, 2)
                        ritmo_medio = act.get('average_speed', 0) # metri al secondo
                        ritmo_min_km = f"{int(1000/ritmo_medio // 60)}'{int(1000/ritmo_medio % 60):02d}\"" if ritmo_medio > 0 else "N/A"
                        
                        strava_effettivo = f"**{act['name']}**<br>Distanza: {distanza} km<br>Ritmo: {ritmo_min_km}/km"
                    else:
                        strava_effettivo = "Nessuna attività registrata oggi su Strava."
                else:
                    strava_effettivo = "⚠️ Errore di connessione a Strava. Controlla il Token."
            else:
                strava_effettivo = "Inserisci il Token Strava nel menu laterale per vedere i risultati."

            st.markdown(f"""
            <div class="card-actual">
                <div class="card-title">🟠 Eseguito su Strava</div>
                <p style="font-size: 1.1rem; margin:0;">{strava_effettivo}</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br><hr><br>", unsafe_allow_html=True)

        # --- SEZIONE: STORICO DIVISO PER MESI ---
        st.subheader("📚 Storico Allenamenti")
        if dati_completi:
            # Crea dei tab dinamici per ogni mese trovato nell'Excel
            tabs = st.tabs(list(dati_completi.keys()))
            
            for i, (nome_mese, df_mese) in enumerate(dati_completi.items()):
                with tabs[i]:
                    st.dataframe(df_mese, use_container_width=True, hide_index=True)
        else:
            st.info("Non sono riuscito a estrarre i dati. Assicurati che le date e gli allenamenti siano allineati orizzontalmente.")

    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file: {e}")

else:
    # Schermata di benvenuto quando non c'è nessun file
    st.info("👈 Apri il menu laterale e carica il file Excel contenente i mesi per iniziare.")
