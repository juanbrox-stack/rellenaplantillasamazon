import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Amazon Pro Filler", layout="wide")
st.title("🚀 Amazon Template Automator (Modo Flexible)")

# --- CONFIGURACIÓN ---
# (Mantenemos tus diccionarios de VALORES_FIJOS y MAPEO_VARIABLES igual que antes)
# ... [Diccionarios omitidos aquí para brevedad, pero incluidos en la ejecución] ...

def clean_name(text):
    """Limpia nombres técnicos para facilitar el emparejamiento"""
    if not text: return ""
    # Quitamos #1, .value, puntos y espacios, y pasamos a minúsculas
    text = str(text).lower()
    text = re.sub(r'#\d+', '', text)
    text = re.sub(r'\.value|\.unit|\.type', '', text)
    return re.sub(r'[^a-z0-9]', '', text)

col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Ejecutar y Rellenar"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            wb = load_workbook(amz_template)
            sheet_name = "Template" if "Template" in wb.sheetnames else wb.sheetnames[1]
            ws = wb[sheet_name]
            
            # --- MAPEO INTELIGENTE ---
            # Leemos fila 4 y guardamos tanto el nombre real como el "limpio"
            amazon_cols_raw = {} 
            amazon_cols_clean = {}
            
            for i, cell in enumerate(ws[4], 1):
                if cell.value:
                    raw_val = str(cell.value).strip()
                    amazon_cols_raw[raw_val] = i
                    amazon_cols_clean[clean_name(raw_val)] = i
            
            st.write("### 🔍 Informe de Emparejamiento")
            
            def get_col_index(target_name):
                # 1. Intento exacto
                if target_name in amazon_cols_raw: return amazon_cols_raw[target_name]
                # 2. Intento limpio
                target_clean = clean_name(target_name)
                if target_clean in amazon_cols_clean: return amazon_cols_clean[target_clean]
                return None

            # --- ESCRITURA ---
            rows_filled = 0
            mapped_log = []

            for i, row_pim in df_pim.iterrows():
                current_row = i + 7 
                
                # Procesar Variables
                for amz_key, pim_col in MAPEO_VARIABLES.items():
                    idx = get_col_index(amz_key)
                    if idx and pim_col in df_pim.columns:
                        ws.cell(row=current_row, column=idx).value = row_pim[pim_col]
                        if i == 0: mapped_log.append(f"✅ {amz_key} -> Col {idx}")
                
                # Procesar Fijos
                for amz_key, val in VALORES_FIJOS.items():
                    idx = get_col_index(amz_key)
                    if idx:
                        ws.cell(row=current_row, column=idx).value = val
                
                rows_filled += 1

            with st.expander("Ver detalles de columnas emparejadas"):
                st.write(mapped_log)

            # --- DESCARGA ---
            output = BytesIO()
            wb.save(output)
            st.success(f"✅ Se han procesado {rows_filled} productos.")
            st.download_button("📥 Descargar Archivo Final", output.getvalue(), "Amazon_Relleno.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")