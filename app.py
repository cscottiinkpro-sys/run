import streamlit as st
import pandas as pd
import requests
import datetime

# Configurazione della pagina
st.set_page_config(page_title="Luca Tassatori Coach", page_icon="🏃‍♂️", layout="centered")

# Intestazione personalizzata
st.title("🏃‍♂️ Luca Tassatori Coach")
st.markdown("*“La fatica di oggi sarà il ritmo di gara di domani.”*")
st.markdown("---")

# Sidebar per caricare il file e configurare Strava
st.sidebar.header("📁 Pannello di Controllo")
uploaded_file = st.sidebar.file_uploader("Carica il file Excel del Coach", type=["xlsx", "xls"])

st.sidebar.markdown("---")
st.sidebar.header("🔗 Integrazione Strava")
strava_token = st.sidebar.text_input("Inserisci Strava Access Token", type="password", help="Ottieni il token temporaneo dalle impostazioni API del tuo account Strava.")

# Sezione Principale: Lettura Excel
if uploaded_file is not None:
    try:
        # Legge il file excel senza intestazione fissa per gestire la struttura orizzontale
        df = pd.read_excel(uploaded_file, header=None)
        
        # Nello screenshot del coach: riga 4 (indice 3) = Date, riga 6 (indice 5) = Allenamenti
        row_dates = df.iloc[3].values
        row_workouts = df.iloc[5].values
        
        workouts_data = []
        for i in range(1, len(row_dates)):
            d = str(row_dates[i]).strip()
            w = str(row_workouts[i]).strip()
            if d and d != "nan" and w and w != "nan":
                workouts_data.append({"Data": d, "Allenamento": w})
        
        if workouts_data:
            st.subheader("📅 Tabella Allenamenti Programmati")
            df_work = pd.DataFrame(workouts_data)
            
            # Mostra l'allenamento in evidenza (il primo disponibile o odierno)
            st.success(f"🔥 **Prossimo Allenamento ({workouts_data[0]['Data']}):** {workouts_data[0]['Allenamento']}")
            
            with st.expander("Visualizza l'intera tabella mensile"):
                st.dataframe(df_work, use_container_width=True)
        else:
            st.warning("⚠️ Non ho trovato date e allenamenti nelle celle previste. Controlla il formato del file.")
            
    except Exception as e:
        st.error(f"Errore nella lettura del file Excel: {e}")
else:
    st.info("👈 Carica il file Excel del coach dal menu laterale per iniziare.")

st.markdown("---")

# Sezione Strava: Confronto Allenamento vs Effettivo
st.subheader("📊 Ultime Attività da Strava")

if strava_token:
    headers = {'Authorization': f'Bearer {strava_token}'}
    try:
        # Chiamata alle API di Strava per recuperare le ultime 5 attività
        response = requests.get('https://www.strava.com/api/v3/athlete/activities?per_page=5', headers=headers)
        
        if response.status_code == 200:
            activities = response.json()
            if activities:
                for act in activities:
                    name = act.get('name')
                    distance_km = round(act.get('distance', 0) / 1000, 2)
                    moving_time_min = round(act.get('moving_time', 0) / 60, 1)
                    date_str = act.get('start_date_local', '')[:10]
                    
                    st.markdown(f"""
                    - **{date_str}**: {name} 
                      * 📏 **Distanza:** {distance_km} km
                      * ⏱️ **Tempo:** {moving_time_min} min
                    """)
            else:
                st.info("Nessuna attività recente trovata su Strava.")
        else:
            st.error("Token Strava non valido o scaduto. Verifica i permessi nelle impostazioni Strava.")
    except Exception as e:
        st.error(f- "Impossibile connettersi a Strava: {e}")
else:
    st.warning("Inserisci il tuo Access Token di Strava nella barra laterale per sincronizzare i tuoi allenamenti effettivi e confrontarli con quelli di Luca Tassatori.")