import streamlit as st
import pandas as pd
from datetime import datetime
from garminconnect import Garmin
import os
import random

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
    hr.garmin-divider { margin: 10px 0; border: 0; border-top: 1px solid #bbf7d0; }
    [data-testid="stDataFrame"] { width: 100%; }
    </style>
""", unsafe_allow_html=True)

# 2. CREDENZIALI GARMIN FISSE E SALVATE
GARMIN_EMAIL = "scocla@hotmail.it"
GARMIN_PWD = "Ciccio1994"
EXCEL_FILE_PATH = "storico_salvato.xlsx"

# Frasi motivazionali per il runner
frasi_motivazionali = [
    "“La fatica di oggi sarà il ritmo di gara di domani.”",
    "“Non correre solo con le gambe, corri con la testa e con il cuore.”",
    "“Il dolore è temporaneo, la gloria di aver finito è per sempre.”",
    "“Ogni passo avanti è un passo più vicino al tuo obiettivo.”",
    "“Se vuoi vincere qualcosa, corri 100 metri. Se vuoi goderti un'altra vita, corri la maratona.”",
    "“Il miracolo non è che ho terminato. Il miracolo è che ho avuto il coraggio di iniziare.”"
]
frase_del_giorno = frasi_motivazionali[datetime.today().timetuple().tm_yday % len(frasi_motivazionali)]

# 3. DATE E CALENDARIO
oggi = datetime.today()
giorno_oggi =oggi.day  
mese_oggi = oggi.month  

mesi_completi = {1:'gennaio', 2:'febbraio', 3:'marzo', 4:'aprile', 5:'maggio', 6:'giugno', 7:'luglio', 8:'agosto', 9:'settembre', 10:'ottobre', 11:'novembre', 12:'dicembre'}
mesi_it = {1:'gen', 2:'feb', 3:'mar', 4:'apr', 5:'mag', 6:'giu', 7:'lug', 8:'ago', 9:'set', 10:'ott', 11:'nov', 12:'dic'}

foglio_default_target = f"{mesi_completi[mese_oggi]}{str(oggi.year)[-2:]}"

st.markdown("<div class='home-logo'>👟</div>", unsafe_allow_html=True)
st.markdown("<h1 class='main-title'>Luca Tassarotti Coach</h1>", unsafe_allow_html=True)
st.markdown(f"<div class='motivation-box'>{frase_del_giorno}</div>", unsafe_allow_html=True)

# 4. MENU LATERALE SOLO PER AGGIORNARE L'EXCEL
st.sidebar.title("⚙️ Impostazioni")
uploaded_file = st.sidebar.file_uploader("Aggiorna File Excel", type=["xlsx", "xls"])
if uploaded_file is not None:
    with open(EXCEL_FILE_PATH, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.sidebar.success("File Excel aggiornato e salvato con successo!")

file_da_leggere = None
if os.path.exists(EXCEL_FILE_PATH):
    file_da_leggere = EXCEL_FILE_PATH

# 5. ELABORAZIONE DATI EXCEL
if file_da_leggere:
    try:
        xls = pd.ExcelFile(file_da_leggere)
        fogli_mesi = xls.sheet_names
        
        indice_default = 0
        for i, nome_foglio in enumerate(fogli_mesi):
            if foglio_default_target in nome_foglio.lower():
                indice_default = i
                break

        dati_completi = {}
        allenamento_oggi = "Riposo o nessun allenamento trovato."
        
        for foglio in fogli_mesi:
            df = pd.read_excel(xls, sheet_name=foglio, header=None)
            mese_data = []
            
            riga_date_idx = -1
            for r_idx in range(len(df)):
                riga_valori = df.iloc[r_idx].values
                for val in riga_valori:
                    val_str = str(val).lower()
                    if f"-{mesi_it[mese_oggi]}" in val_str or f"/{mese_oggi}/" in val_str or "lug" in val_str or "jul" in val_str:
                        riga_date_idx = r_idx
                        break
                if riga_date_idx != -1:
                    break
            
            if riga_date_idx == -1 and len(df) > 3:
                riga_date_idx = 3
            
            riga_workout_idx = riga_date_idx + 2 if (riga_date_idx + 2) < len(df) else riga_date_idx + 1
            
            if riga_date_idx != -1 and riga_workout_idx < len(df):
                row_dates = df.iloc[riga_date_idx].values
                row_workouts = df.iloc[riga_workout_idx].values
                
                for col_idx in range(len(row_dates)):
                    d_raw = row_dates[col_idx]
                    w = str(row_workouts[col_idx]).strip() if col_idx < len(row_workouts) else ""
                    
                    if pd.notna(d_raw):
                        d_str = ""
                        is_oggi = False
                        
                        if isinstance(d_raw, datetime) or type(d_raw).__name__ == 'Timestamp':
                            d_str = f"{d_raw.day}-{mesi_it.get(d_raw.month, '')}"
                            if d_raw.day == giorno_oggi and d_raw.month == mese_oggi:
                                is_oggi = True
                        else:
                            d_str = str(d_raw).strip()
                            d_lower = d_str.lower()
                            if (f"{giorno_oggi}-" in d_lower or f"{giorno_oggi} " in d_lower or d_lower.startswith(str(giorno_oggi))) and ("lug" in d_lower or "jul" in d_lower or str(mese_oggi) in d_lower):
                                is_oggi = True
                        
                        if d_str and d_str != "nan":
                            mese_data.append({"Data": d_str.capitalize(), "Programma": w if w != "nan" else "Riposo"})
                        
                        if is_oggi and w and w != "nan" and w != "":
                            allenamento_oggi = w

            if mese_data:
                dati_completi[foglio] = pd.DataFrame(mese_data)

        # --- MOSTRA ALLENAMENTO PROGRAMMATO (COACH) ---
        st.subheader(f"⚡ Oggi ({oggi.strftime('%d/%m/%Y')})")
        
        st.markdown(f"""
        <div class="card-planned">
            <div class="card-title">🔵 Da Tabella Coach</div>
            <p class="workout-text">{allenamento_oggi}</p>
        </div>
        """, unsafe_allow_html=True)

        # --- MOSTRA TUTTI GLI ALLENAMENTI GARMIN DEL GIORNO ---
        garmin_effettivo = "Sincronizzazione in corso..."
        try:
            client = Garmin(GARMIN_EMAIL, GARMIN_PWD)
            client.login()
            
            oggi_iso = oggi.strftime('%Y-%m-%d')
            attivita = client.get_activities(0, 10) 
            attivita_oggi = [a for a in attivita if a.get('startTimeLocal', '').startswith(oggi_iso)]
            
            if attivita_oggi:
                garmin_effettivo = ""
                for i, act in enumerate(attivita_oggi):
                    distanza = round(act.get('distance', 0) / 1000, 2)
                    durata_min = round(act.get('duration', 0) / 60, 1)
                    velocita_ms = act.get('averageSpeed', 0)
                    
                    passo_str = f"{int((1000/velocita_ms)//60)}'{int((1000/velocita_ms)%60):02d}\"" if velocita_ms > 0 else "N/A"
                    
                    if i > 0:
                        garmin_effettivo += "<hr class='garmin-divider'>"
                        
                    garmin_effettivo += f"<p class='workout-text'><strong>{act.get('activityName', 'Attività')}</strong><br>📏 {distanza} km &nbsp;|&nbsp; ⏱️ {durata_min} min &nbsp;|&nbsp; ⚡ {passo_str}/km</p>"
            else:
                garmin_effettivo = "<p class='workout-text'>Nessuna attività registrata oggi su Garmin.</p>"
        except Exception as e:
            garmin_effettivo = f"<p class='workout-text'>⚠️ Errore di connessione a Garmin. Riprova più tardi.</p>"

        st.markdown(f"""
        <div class="card-actual">
            <div class="card-title">🟢 Svolto su Garmin</div>
            {garmin_effettivo}
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        # --- STORICO (Parte dal mese corrente in automatico) ---
        st.subheader("📚 Storico Allenamenti")
        if dati_completi:
            mese_scelto = st.selectbox("Seleziona periodo:", list(dati_completi.keys()), index=indice_default)
            st.dataframe(dati_completi[mese_scelto], use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Errore tecnico durante la lettura dell'Excel: {e}")

else:
    st.info("👈 Apri il menu laterale in alto a sinistra per caricare l'Excel per la prima volta.")
