import streamlit as st
import pandas as pd
import re
import io
import os
import pickle
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ====================== CONFIGURAZIONE ======================
DB_FILE = "dati_trasporti_salvati.pkl"

def carica_dati_permanenti():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "rb") as f:
                return pickle.load(f)
        except:
            return {"accumulated_rows": [], "imported_files": [], "extras": []}
    return {"accumulated_rows": [], "imported_files": [], "extras": []}

def salva_dati_permanenti():
    with open(DB_FILE, "wb") as f:
        pickle.dump({
            "accumulated_rows": st.session_state.accumulated_rows,
            "imported_files": st.session_state.imported_files,
            "extras": st.session_state.extras
        }, f)

# ====================== PARSING PDF ======================
def parse_nippon_pdf(file) -> pd.DataFrame:
    """Estrae e pulisce i dati dalle Liste di Carico Nippon Express"""
    rows = []
    
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                df_page = pd.DataFrame(table[1:], columns=table[0])
                df_page = df_page.dropna(how='all').reset_index(drop=True)
                
                # Normalizzazione nomi colonne
                df_page.columns = [re.sub(r'\s+', ' ', str(col).strip().upper()) for col in df_page.columns]
                
                col_map = {
                    'MITTENTE': 'MITTENTE',
                    'DESTINATARIO': 'DESTINATARIO',
                    'COLLI': 'IMBALLI',
                    'IMBALLI': 'IMBALLI',
                    'VOLUME': 'VOLUME',
                    'MC': 'VOLUME',
                    'LOCALITÀ': 'LOCALITÀ',
                    'LOCALITA': 'LOCALITÀ',
                    'LOCALITA\'': 'LOCALITÀ'
                }
                
                for old, new in col_map.items():
                    for col in df_page.columns:
                        if old in col:
                            df_page = df_page.rename(columns={col: new})
                
                # Pulizia VOLUME (gestione doppio punto, MC, ecc.)
                def clean_volume(v):
                    if pd.isna(v):
                        return 0.0
                    s = str(v).strip().upper()
                    s = re.sub(r'[^\d.,]', '', s)          # rimuove MC, SI, NO...
                    s = s.replace(',', '.')
                    s = re.sub(r'\.(\d+)\.(\d+)', r'.\1\2', s)  # 125.20.00 → 125.2000
                    try:
                        return round(float(s), 4)
                    except:
                        return 0.0
                
                if 'VOLUME' in df_page.columns:
                    df_page['VOLUME'] = df_page['VOLUME'].apply(clean_volume)
                
                # Pulizia IMBALLI
                if 'IMBALLI' in df_page.columns:
                    df_page['IMBALLI'] = pd.to_numeric(df_page['IMBALLI'], errors='coerce').fillna(0).astype(int)
                
                # Pulizia nomi
                def clean_name(name):
                    if pd.isna(name):
                        return ""
                    name = str(name).strip()
                    # Rimuove suffissi societari
                    suffixes = ['S.R.L.', 'SRL', 'S.P.A.', 'SPA', 'S.N.C.', 'SNC', 'GROUP', 'FURNITURE', 'ITALIA', 'S.A.S.', 'SAS']
                    for s in suffixes:
                        name = re.sub(rf'\b{s}\b', '', name, flags=re.IGNORECASE)
                    name = re.sub(r'\s+', ' ', name).strip()
                    return name[:30]
                
                for col in ['MITTENTE', 'DESTINATARIO']:
                    if col in df_page.columns:
                        df_page[col] = df_page[col].apply(clean_name)
                
                rows.extend(df_page.to_dict('records'))
    
    df = pd.DataFrame(rows)
    # Colonne essenziali
    for col in ['MITTENTE', 'DESTINATARIO', 'IMBALLI', 'VOLUME', 'LOCALITÀ']:
        if col not in df.columns:
            df[col] = None if col in ['MITTENTE', 'DESTINATARIO', 'LOCALITÀ'] else 0
    
    return df[['MITTENTE', 'DESTINATARIO', 'IMBALLI', 'VOLUME', 'LOCALITÀ']]

# ====================== STREAMLIT APP ======================
st.set_page_config(page_title="DMI LogiSmart", layout="wide")
st.title("🚚 DMI LogiSmart - Gestione Liste di Carico")

# Inizializzazione sessione
if 'initialized' not in st.session_state:
    dati = carica_dati_permanenti()
    st.session_state.accumulated_rows = dati["accumulated_rows"]
    st.session_state.imported_files = dati["imported_files"]
    st.session_state.extras = dati["extras"]
    st.session_state.initialized = True

# Sidebar
st.sidebar.header("⚙️ Controlli")
if st.session_state.accumulated_rows:
    st.sidebar.info(f"Righe totali: {len(st.session_state.accumulated_rows)}")
    if st.sidebar.button("🗑️ Svuota Archivio (Nuova Settimana)"):
        st.session_state.accumulated_rows = []
        st.session_state.imported_files = []
        st.session_state.extras = []
        salva_dati_permanenti()
        st.rerun()

# Upload
file_caricati = st.file_uploader(
    "Trascina qui i PDF delle Liste di Carico (multiplo)",
    type="pdf",
    accept_multiple_files=True
)

if file_caricati:
    file_da_importare = [f for f in file_caricati if f.name not in st.session_state.imported_files]
    
    if file_da_importare and st.button("📥 Elabora e Aggiungi all'Archivio"):
        progress = st.progress(0)
        
        for i, file in enumerate(file_caricati):
            try:
                nome_autista = re.sub(r'\.pdf$', '', file.name, flags=re.IGNORECASE).strip().upper()
                
                df_file = parse_nippon_pdf(file)
                
                if len(df_file) == 0:
                    st.error(f"Impossibile estrarre dati da {file.name}")
                    continue
                
                df_file['AUTISTA'] = nome_autista
                for _, row in df_file.iterrows():
                    st.session_state.accumulated_rows.append(row.to_dict())
                
                st.session_state.imported_files.append(file.name)
                st.success(f"✅ {file.name} → {len(df_file)} righe")
                
            except Exception as e:
                st.error(f"Errore su {file.name}: {str(e)}")
            
            progress.progress((i + 1) / len(file_caricati))
        
        salva_dati_permanenti()
        st.rerun()

# ====================== ELABORAZIONE ======================
if st.session_state.accumulated_rows:
    df = pd.DataFrame(st.session_state.accumulated_rows)
    
    # Aggregazione
    df_agg = df.groupby(['AUTISTA', 'MITTENTE', 'DESTINATARIO', 'LOCALITÀ'], as_index=False).agg({
        'IMBALLI': 'sum',
        'VOLUME': 'sum'
    }).round({'VOLUME': 4})
    
    # Logica Minime
    vol_per_dest = df.groupby('DESTINATARIO')['VOLUME'].sum()
    destinatari_grandi = vol_per_dest[vol_per_dest >= 1.0].index.tolist()
    
    df_agg['CONSEGNA_MINIMA'] = df_agg.apply(
        lambda r: "SI" if (r['VOLUME'] < 1.0 and r['DESTINATARIO'] not in destinatari_grandi) else "NO",
        axis=1
    )
    
    # Editor interattivo
    st.header("📋 Dati Elaborati - Modifica se necessario")
    edited_df = st.data_editor(
        df_agg[['AUTISTA', 'MITTENTE', 'DESTINATARIO', 'LOCALITÀ', 'IMBALLI', 'VOLUME', 'CONSEGNA_MINIMA']],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "IMBALLI": st.column_config.NumberColumn(format="%d"),
            "VOLUME": st.column_config.NumberColumn(format="%.4f"),
        }
    )
    
    if st.button("💾 Genera Report Excel"):
        # Preparazione per Excel
        buffer = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Report Consegne"
        
        # Stili
        title_font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
        title_fill = PatternFill('solid', start_color='1F4E78')
        header_font = Font(name='Calibri', size=11, bold=True)
        header_fill = PatternFill('solid', start_color='D9E1F2')
        red_fill = PatternFill('solid', start_color='FFC7CE')
        red_font = Font(name='Calibri', color='9C0006', bold=True)
        
        row = 1
        
        for autista in edited_df['AUTISTA'].unique():
            df_aut = edited_df[edited_df['AUTISTA'] == autista]
            
            # Titolo autista
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            cell = ws.cell(row=row, column=1, value=f"SPEDIZIONE AUTISTA: {autista}")
            cell.font = title_font
            cell.fill = title_fill
            cell.alignment = Alignment(horizontal='center')
            row += 1
            
            # Intestazioni
            headers = ["MITTENTE", "DESTINATARIO", "COLLI", "METRI CUBI", "LOCALITÀ"]
            for col, h in enumerate(headers, 1):
                c = ws.cell(row=row, column=col, value=h)
                c.font = header_font
                c.fill = header_fill
            row += 1
            
            # Dati
            for _, r in df_aut.iterrows():
                ws.cell(row=row, column=1, value=r['MITTENTE'])
                ws.cell(row=row, column=2, value=r['DESTINATARIO'])
                ws.cell(row=row, column=3, value=int(r['IMBALLI']))
                ws.cell(row=row, column=4, value=float(r['VOLUME']))
                ws.cell(row=row, column=5, value=r['LOCALITÀ'])
                
                if r['CONSEGNA_MINIMA'] == "SI":
                    ws.cell(row=row, column=4).fill = red_fill
                    ws.cell(row=row, column=4).font = red_font
                
                row += 1
            row += 2
        
        # Auto-width colonne
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 18
        
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📥 Scarica Report Excel Professionale",
            data=buffer,
            file_name=f"LogiSmart_Report_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Carica i primi PDF per iniziare la settimana.")
