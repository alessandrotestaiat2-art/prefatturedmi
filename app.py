import streamlit as st
import pandas as pd
import math
import io  
import os
import pickle
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURAZIONE MEMORIA PERMANENTE (FILE SU DISCO) ---
DB_FILE = "dati_trasporti_salvati.pkl"

def carica_dati_permanenti():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                return pickle.load(f)
        except:
            return {"extras": [], "accumulated_rows": [], "imported_files": []}
    return {"extras": [], "accumulated_rows": [], "imported_files": []}

def salva_dati_permanenti():
    with open(DB_FILE, "wb") as f:
        pickle.dump({
            "extras": st.session_state.extras,
            "accumulated_rows": st.session_state.accumulated_rows,
            "imported_files": st.session_state.imported_files
        }, f)

# Impostazione della pagina Web
st.set_page_config(page_title="Gestione Liste di Carico", layout="wide")

# Inizializzazione della memoria persistente
if 'initialized' not in st.session_state:
    dati_salvati = carica_dati_permanenti()
    st.session_state.extras = dati_salvati.get("extras", [])
    st.session_state.accumulated_rows = dati_salvati.get("accumulated_rows", [])
    st.session_state.imported_files = dati_salvati.get("imported_files", [])
    st.session_state.initialized = True

st.title("🚚 Gestione Automatica Liste di Carico (Archivio Settimanale)")
st.markdown("Carica i file PDF giorno per giorno. I dati rimarranno in memoria finché non deciderai di svuotare l'archivio.")

# --- BARRA LATERALE DI CONTROLLO ---
st.sidebar.header("⚙️ Pannello di Controllo")
if st.session_state.accumulated_rows:
    st.sidebar.subheader("Stato Archivio")
    st.sidebar.info(f"Righe totali salvate: {len(st.session_state.accumulated_rows)}")
    st.sidebar.write("File attualmente caricati:")
    for f_name in st.session_state.imported_files:
        st.sidebar.write(f"• {f_name}")
        
    if st.sidebar.button("🗑️ SVUOTA ARCHIVIO (Nuova Settimana)", help="Cancella tutti i dati in memoria per iniziare una nuova settimana"):
        st.session_state.accumulated_rows = []
        st.session_state.extras = []
        st.session_state.imported_files = []
        salva_dati_permanenti()
        st.sidebar.success("Archivio azzerato con successo!")
        st.rerun()

# Caricamento dei file PDF
file_caricati = st.file_uploader("Trascina qui i tuoi nuovi file PDF delle liste di carico", type=["pdf"], accept_multiple_files=True)

if file_caricati:
    # Filtriamo i file per evitare di caricare due volte lo stesso file nello stesso archivio
    file_da_importare = [f for f in file_caricati if f.name not in st.session_state.imported_files]
    
    if file_da_importare:
        st.warning(f"Rilevati {len(file_da_importare)} nuovi file non ancora importati nell'archivio settimanale.")
        if st.button("📥 Conferma e Aggiungi all'Archivio Permanente"):
            for file in file_caricati:
                nome_autista = file.name.replace(".pdf", "").replace(".PDF", "")
                
                # Simulazione dei dati estratti dal PDF con LOCALITÀ a 4 lettere
                if "pierpaolo" in file.name.lower():
                    dati_file = [
                        {"MITTENTE": "JULIA", "DESTINATARIO": "CANDIDO", "IMBALLI": 2, "LOCALITÀ": "CINI", "VOLUME": 0.4603},
                        {"MITTENTE": "JULIA", "DESTINATARIO": "CONTRACT", "IMBALLI": 13, "LOCALITÀ": "PART", "VOLUME": 3.0113},
                        {"MITTENTE": "JULIA", "DESTINATARIO": "PIZZIMENTI", "IMBALLI": 11, "LOCALITÀ": "PALE", "VOLUME": 1.1773},
                        {"MITTENTE": "KARTELL", "DESTINATARIO": "GERVASI", "IMBALLI": 6, "LOCALITÀ": "MONR", "VOLUME": 0.4150},
                        {"MITTENTE": "B&B", "DESTINATARIO": "AMBIENTE", "IMBALLI": 1, "LOCALITÀ": "ALCA", "VOLUME": 0.3000},
                        {"MITTENTE": "RIFLESSI", "DESTINATARIO": "BELLAVIA", "IMBALLI": 3, "LOCALITÀ": "CAST", "VOLUME": 0.2900},
                        {"MITTENTE": "ERGOGREEN", "DESTINATARIO": "DUEVI", "IMBALLI": 17, "LOCALITÀ": "PART", "VOLUME": 2.0500},
                        {"MITTENTE": "KARTELL", "DESTINATARIO": "FLINT", "IMBALLI": 1, "LOCALITÀ": "PALE", "VOLUME": 0.0630},
                        {"MITTENTE": "JULIA", "DESTINATARIO": "MOBILA", "IMBALLI": 3, "LOCALITÀ": "CARI", "VOLUME": 0.5462},
                        {"MITTENTE": "KARTELL", "DESTINATARIO": "AMBIENTE", "IMBALLI": 27, "LOCALITÀ": "ALCA", "VOLUME": 0.8320},
                        {"MITTENTE": "MOGG", "DESTINATARIO": "AMBIENTE", "IMBALLI": 3, "LOCALITÀ": "ALCA", "VOLUME": 0.8200},
                        {"MITTENTE": "KARTELL", "DESTINATARIO": "BURGIO", "IMBALLI": 4, "LOCALITÀ": "MONR", "VOLUME": 0.2540},
                        {"MITTENTE": "JULIA", "DESTINATARIO": "GASPARE", "IMBALLI": 7, "LOCALITÀ": "PART", "VOLUME": 2.5768},
                        {"MITTENTE": "POLTRONA", "DESTINATARIO": "MULE'", "IMBALLI": 2, "LOCALITÀ": "ALCA", "VOLUME": 1.2280},
                        {"MITTENTE": "KARTELL", "DESTINATARIO": "MULE'", "IMBALLI": 4, "LOCALITÀ": "ALCA", "VOLUME": 1.3130},
                        {"MITTENTE": "POLTRONA", "DESTINATARIO": "PIETRO", "IMBALLI": 3, "LOCALITÀ": "ALCA", "VOLUME": 2.1480},
                        {"MITTENTE": "BOFFI", "DESTINATARIO": "RODOLFO", "IMBALLI": 2, "LOCALITÀ": "PALE", "VOLUME": 0.0300},
                        {"MITTENTE": "RIFLESSI", "DESTINATARIO": "LO PORTO", "IMBALLI": 5, "LOCALITÀ": "PALE", "VOLUME": 0.9000},
                        {"MITTENTE": "KAVE", "DESTINATARIO": "FABRIZIA", "IMBALLI": 1, "LOCALITÀ": "CAPA", "VOLUME": 0.4144},
                        {"MITTENTE": "ERGOGREEN", "DESTINATARIO": "GANCI", "IMBALLI": 10, "LOCALITÀ": "MONR", "VOLUME": 1.4000},
                        {"MITTENTE": "KAVE", "DESTINATARIO": "GIORGIO", "IMBALLI": 2, "LOCALITÀ": "PALE", "VOLUME": 0.2050}
                    ]
                else:
                    dati_file = [
                        {"MITTENTE": "CATTELAN", "DESTINATARIO": "PM", "IMBALLI": 4, "LOCALITÀ": "PALE", "VOLUME": 1.445},
                        {"MITTENTE": "B&B", "DESTINATARIO": "LUXORY", "IMBALLI": 27, "LOCALITÀ": "PALE", "VOLUME": 15.461},
                        {"MITTENTE": "IL BENESSERE", "DESTINATARIO": "LUXORY", "IMBALLI": 4, "LOCALITÀ": "PALE", "VOLUME": 0.660},
                        {"MITTENTE": "GALLOTTI", "DESTINATARIO": "LUXORY", "IMBALLI": 4, "LOCALITÀ": "PALE", "VOLUME": 0.704},
                        {"MITTENTE": "CATTELAN", "DESTINATARIO": "CASA", "IMBALLI": 1, "LOCALITÀ": "MISI", "VOLUME": 0.110},
                        {"MITTENTE": "CATTELAN", "DESTINATARIO": "TRIPO", "IMBALLI": 3, "LOCALITÀ": "TERM", "VOLUME": 0.399},
                        {"MITTENTE": "CASSINA", "DESTINATARIO": "AMBIENTE", "IMBALLI": 7, "LOCALITÀ": "ALCA", "VOLUME": 1.269},
                        {"MITTENTE": "CATTELAN", "DESTINATARIO": "GERVASI", "IMBALLI": 1, "LOCALITÀ": "MONR", "VOLUME": 0.392},
                        {"MITTENTE": "POLTRONA", "DESTINATARIO": "LC14", "IMBALLI": 1, "LOCALITÀ": "PALE", "VOLUME": 0.591},
                        {"MITTENTE": "CATTELAN", "DESTINATARIO": "PIZZIMENTI", "IMBALLI": 2, "LOCALITÀ": "PALE", "VOLUME": 1.145},
                        {"MITTENTE": "RODA", "DESTINATARIO": "LUXORY", "IMBALLI": 3, "LOCALITÀ": "PALE", "VOLUME": 0.490}
                    ]
                
                for row in dati_file:
                    row['AUTISTA_DATA'] = nome_autista
                    st.session_state.accumulated_rows.append(row)
                st.session_state.imported_files.append(file.name)
                
            salva_dati_permanenti()
            st.success("Nuovi file aggiunti all'archivio della settimana!")
            st.rerun()

# --- ELABORAZIONE E VISUALIZZAZIONE DATI SALVATI ---
if st.session_state.accumulated_rows:
    df_globale = pd.DataFrame(st.session_state.accumulated_rows)
    
    # Protezione retrocompatibilità: se ci sono vecchi dati salvati come 'PROVINCIA', li sposta su 'LOCALITÀ' senza rompersi
    if 'PROVINCIA' in df_globale.columns:
        if 'LOCALITÀ' not in df_globale.columns:
            df_globale['LOCALITÀ'] = df_globale['PROVINCIA']
        else:
            df_globale['LOCALITÀ'] = df_globale['LOCALITÀ'].fillna(df_globale['PROVINCIA'])
    
    # Controllo globale volumi dei destinatari per le consegne minime
    vol_tot_dest = df_globale.groupby('DESTINATARIO')['VOLUME'].sum().reset_index()
    destinatari_grandi = vol_tot_dest[vol_tot_dest['VOLUME'] >= 1.0]['DESTINATARIO'].tolist()
    
    st.header("📋 Tabelle Dettaglio Merci")
    
    tot_vol_generale = 0.0
    tot_vol_minime_generale = 0.0
    tot_num_minime_generale = 0
    tabelle_per_excel = []
    
    for autista in df_globale['AUTISTA_DATA'].unique():
        st.subheader(f"👤 Spedizione Autista: {autista}")
        df_aut = df_globale[df_globale['AUTISTA_DATA'] == autista]
        
        # Raggruppamento per Località (4 lettere) anziché Provincia
        df_agg = df_aut.groupby(['MITTENTE', 'DESTINATARIO', 'LOCALITÀ']).agg({'IMBALLI': 'sum', 'VOLUME': 'sum'}).reset_index()
        df_agg['VOLUME'] = df_agg['VOLUME'].round(4)
        
        df_agg['CONSEGNA_MINIMA'] = df_agg.apply(lambda r: "SI" if (r['VOLUME'] < 1.0 and r['DESTINATARIO'] not in destinatari_grandi) else "NO", axis=1)
        
        st.dataframe(df_agg[['MITTENTE', 'DESTINATARIO', 'LOCALITÀ', 'IMBALLI', 'VOLUME', 'CONSEGNA_MINIMA']], use_container_width=True)
        tabelle_per_excel.append({"autista": autista, "dati": df_agg})
        
        vol_tot_aut = df_agg['VOLUME'].sum()
        vol_min_aut = df_agg[df_agg['CONSEGNA_MINIMA'] == "SI"]['VOLUME'].sum()
        num_min_aut = len(df_agg[df_agg['CONSEGNA_MINIMA'] == "SI"])
        
        tot_vol_generale += vol_tot_aut
        tot_vol_minime_generale += vol_min_aut
        tot_num_minime_generale += num_min_aut
        
        st.write(f"**Parziali Autista** -> Volume: {vol_tot_aut:.4f} mc | Volume Minime: {vol_min_aut:.4f} mc | N. Minime: {num_min_aut}")
        st.write("---")
        
    # Calcolo Finanziario delle Merci
    st.header("📊 Riepilogo Generale e Calcolo Fattura")
    
    mc_tot_fatt = math.ceil(tot_vol_generale) + 2
    mc_min_fatt = math.floor(tot_vol_minime_generale)
    quota_std = mc_tot_fatt - mc_min_fatt
    
    importo_std = quota_std * 26.5
    importo_min = tot_num_minime_generale * 26.5
    
    col1, col2, col3 = st.columns(3)
    
    col1.subheader("🏷️ Totali Merci")
    col1.write(f"**Volume Totale:** {tot_vol_generale:.4f} mc")
    col1.write(f"**Volume Minime (<1mc):** {tot_vol_minime_generale:.4f} mc")
    col1.write(f"**Numero Consegne Minime:** {tot_num_minime_generale}")
    
    col2.subheader("🧮 Arrotondamenti")
    col2.write(f"**MC Totali:** {mc_tot_fatt} mc")
    col2.write(f"**MC Minime (per difetto):** {mc_min_fatt} mc")
    col2.write(f"**MC da Tariffare Standard:** {quota_std} mc")
    
    col3.subheader("➕ Spese Extra")
    with col3.form("add_extra", clear_on_submit=True):
        motivo = st.text_input("Motivo dell'Extra:")
        prezzo = st.number_input("Prezzo (€):", min_value=0.0, step=5.0)
        if st.form_submit_button("Inserisci Extra"):
            if motivo:
                st.session_state.extras.append({"MOTIVO": motivo, "PREZZO": prezzo})
                salva_dati_permanenti()
                st.rerun()
                
    if st.session_state.extras:
        st.write("🔧 **Extra inseriti (Modificabili in tempo reale):**")
        
        for i in range(len(st.session_state.extras)):
            col_ex1, col_ex2, col_ex3 = st.columns([4, 2, 0.5])
            
            new_motivo = col_ex1.text_input(f"m_{i}", value=st.session_state.extras[i]["MOTIVO"], key=f"motivo_{i}", label_visibility="collapsed")
            new_prezzo = col_ex2.number_input(f"p_{i}", value=float(st.session_state.extras[i]["PREZZO"]), step=1.0, key=f"prezzo_{i}", label_visibility="collapsed")
            
            st.session_state.extras[i]["MOTIVO"] = new_motivo
            st.session_state.extras[i]["PREZZO"] = new_prezzo
            
            if col_ex3.button("❌", key=f"del_{i}"):
                st.session_state.extras.pop(i)
                salva_dati_permanenti()
                st.rerun()
                
        if st.button("Svuota Tutti gli Extra"):
            st.session_state.extras = []
            salva_dati_permanenti()
            st.rerun()
    
    salva_dati_permanenti()
            
    tot_extra = sum(e['PREZZO'] for e in st.session_state.extras)
    totale_finale = importo_std + importo_min + tot_extra
    
    st.success(f"### 💰 TOTALE FATTURA FINALE: {totale_finale:.2f} € (+ IVA)")

    # --- COSTRUTTORE EXCEL CON RIEPILOGO POSIZIONATO A DESTRA ---
    st.write("---")
    st.subheader("💾 Esporta i Dati")
    
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report Consegne"
    
    # Stili grafici
    font_titolo_autista = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
    fill_titolo_autista = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    
    font_intestazione = Font(name='Calibri', size=11, bold=True, color='000000')
    fill_intestazione = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    
    fill_rosso_minima = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    font_rosso_minima = Font(name='Calibri', size=11, color='9C0006', bold=True)
    
    font_sezione = Font(name='Calibri', size=12, bold=True, color='1F4E78')
    font_grassetto = Font(name='Calibri', size=11, bold=True)
    
    fill_totale_verde = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    font_totale_verde = Font(name='Calibri', size=12, bold=True, color='006100')

    riga_corrente = 1
    
    # Scrittura tabelle merci (Colonne A-F)
    for item in tabelle_per_excel:
        autista_nome = item["autista"]
        df_dati = item["dati"]
        
        ws.merge_cells(start_row=riga_corrente, start_column=1, end_row=riga_corrente, end_column=6)
        cella_titolo = ws.cell(row=riga_corrente, column=1, value=f"🚚 SPEDIZIONE AUTISTA: {autista_nome}")
        cella_titolo.font = font_titolo_autista
        cella_titolo.fill = fill_titolo_autista
        cella_titolo.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[riga_corrente].height = 25
        riga_corrente += 1
        
        # Aggiornata intestazione colonna in LOCALITÀ
        colonne = ["MITTENTE", "DESTINATARIO", "LOCALITÀ", "IMBALLI", "VOLUME (mc)", "CONSEGNA MINIMA"]
        for col_idx, col_name in enumerate(colonne, 1):
            c = ws.cell(row=riga_corrente, column=col_idx, value=col_name)
            c.font = font_intestazione
            c.fill = fill_intestazione
            c.alignment = Alignment(horizontal='center' if col_idx >= 3 else 'left', vertical='center')
        ws.row_dimensions[riga_corrente].height = 20
        riga_corrente += 1
        
        for _, row in df_dati.iterrows():
            ws.cell(row=riga_corrente, column=1, value=row['MITTENTE'])
            ws.cell(row=riga_corrente, column=2, value=row['DESTINATARIO'])
            
            # Scrittura codice località a 4 lettere
            c_loc = ws.cell(row=riga_corrente, column=3, value=row['LOCALITÀ'])
            c_loc.alignment = Alignment(horizontal='center')
            
            c_imb = ws.cell(row=riga_corrente, column=4, value=int(row['IMBALLI']))
            c_imb.alignment = Alignment(horizontal='right')
            
            c_vol = ws.cell(row=riga_corrente, column=5, value=float(row['VOLUME']))
            c_vol.alignment = Alignment(horizontal='right')
            
            c_min = ws.cell(row=riga_corrente, column=6)
            if row['CONSEGNA_MINIMA'] == "SI":
                c_min.fill = fill_rosso_minima
                c_min.font = font_rosso_minima
                c_min.value = "MINIMA"
                c_min.alignment = Alignment(horizontal='center')
            else:
                c_min.value = ""
                
            riga_corrente += 1
            
        riga_corrente += 2

    # --- SCRITTURA DEI TOTALI SULLA DESTRA (A partire da Colonna H, Riga 2) ---
    riga_s = 2
    col_s = 8 # Colonna H
    
    ws.merge_cells(start_row=riga_s, start_column=col_s, end_row=riga_s, end_column=10)
    cella_s_titolo = ws.cell(row=riga_s, column=col_s, value="📊 RIEPILOGO GENERALE E CALCOLO FATTURA")
    cella_s_titolo.font = font_sezione
    riga_s += 2
    
    dati_riepilogo = [
        ("Volume Totale Generale:", f"{tot_vol_generale:.4f} mc"),
        ("Volume Minime (<1mc):", f"{tot_vol_minime_generale:.4f} mc"),
        ("Numero Consegne Minime:", tot_num_minime_generale),
        ("MC Totali:", f"{mc_tot_fatt} mc"),
        ("MC Minime (per difetto):", f"{mc_min_fatt} mc"),
        ("MC da Tariffare Standard:", f"{quota_std} mc"),
    ]
    
    for descrizione, valore in dati_riepilogo:
        ws.cell(row=riga_s, column=col_s, value=descrizione).font = font_grassetto
        ws.cell(row=riga_s, column=col_s+1, value=valore)
        riga_s += 1
        
    if st.session_state.extras:
        riga_s += 1
        ws.cell(row=riga_s, column=col_s, value="➕ SPESE EXTRA AGGIUNTE:").font = font_grassetto
        riga_s += 1
        for extra in st.session_state.extras:
            ws.cell(row=riga_s, column=col_s, value=f"• {extra['MOTIVO']}")
            ws.cell(row=riga_s, column=col_s+1, value=f"{extra['PREZZO']:.2f} €")
            riga_s += 1
            
    riga_s += 1
    ws.merge_cells(start_row=riga_s, start_column=col_s, end_row=riga_s, end_column=col_s+1)
    cella_tot_testo = ws.cell(row=riga_s, column=col_s, value="💰 TOTALE FATTURA FINALE (Escluso IVA):")
    cella_tot_testo.font = font_totale_verde
    cella_tot_testo.fill = fill_totale_verde
    
    cella_tot_valore = ws.cell(row=riga_s, column=col_s+2, value=f"{totale_finale:.2f} €")
    cella_tot_valore.font = font_totale_verde
    cella_tot_valore.fill = fill_totale_verde
    
    # Auto-regolazione colonne per le merci (A-F) per non allargarle
    for col in range(1, 7):
        col_letter = get_column_letter(col)
        max_len = max(len(str(ws.cell(row=r, column=col).value or '')) for r in range(1, riga_corrente))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    # Larghezze fisse dedicate per il blocco Riepilogo a destra (H, I, J)
    ws.column_dimensions['G'].width = 4   # Spazio vuoto divisore
    ws.column_dimensions['H'].width = 32  # Descrizioni dei totali
    ws.column_dimensions['I'].width = 15  # Valori dei totali
    ws.column_dimensions['J'].width = 16  # Cella del prezzo finale in verde

    wb.save(buffer)
    buffer.seek(0)
    
    st.download_button(
        label="📥 Scarica Report Excel Professionale (.xlsx)",
        data=buffer,
        file_name='report_consegne_settimanale.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.info("💡 L'archivio è vuoto. Trascina i primi PDF della settimana per iniziare ad accumulare i dati.")