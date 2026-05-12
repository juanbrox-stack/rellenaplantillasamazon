import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Amazon Pro Filler", layout="wide")

st.title("🚀 Amazon Template Automator")

# --- FUNCIONES DE LIMPIEZA EXTREMA ---
def super_clean(text):
    if text is None: return ""
    # Elimina todo lo que no sea letra o número y pasa a minúsculas
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

# --- DICCIONARIOS POR DEFECTO ---
VALORES_FIJOS = {
    "brand#1.value": "Cecotec", "external_product_id#1.type": "EAN", 
    "package_level#1.value": "Unit", "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain", "item_weight#1.unit": "Kilograms"
}

MAPEO_INICIAL = {
    "vendor_sku#1.value": "SKU", "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN", "model_number#1.value": "Nombre Producto / Modelo"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Configuración")
final_mapping = {}
with st.sidebar.expander("📝 Editar Mapeo PIM"):
    for amz, pim in MAPEO_INICIAL.items():
        final_mapping[amz] = st.text_input(f"AMZ: {amz}", value=pim)

# --- CARGA ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Procesar"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            wb = load_workbook(amz_template)
            sheet_name = "Template" if "Template" in wb.sheetnames else wb.sheetnames[1]
            ws = wb[sheet_name]
            
            # --- SECCIÓN DE DIAGNÓSTICO VISUAL ---
            st.write("### 🔍 Análisis de la Fila 4 (Amazon)")
            detected_cols = {}
            debug_list = []
            
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=4, column=col_idx).value
                if cell_value:
                    clean_name = super_clean(cell_value)
                    detected_cols[clean_name] = col_idx
                    if col_idx < 15: # Solo mostrar las primeras 15 para no saturar
                        debug_list.append(f"Col {col_idx}: `{cell_value}`")
            
            st.code("\n".join(debug_list), language="text")

            # --- ESCRITURA ---
            count = 0
            for i, row_pim in df_pim.iterrows():
                target_row = i + 7
                
                # Rellenar Variables
                for amz_key, pim_col in final_mapping.items():
                    target_clean = super_clean(amz_key)
                    if target_clean in detected_cols and pim_col in df_pim.columns:
                        ws.cell(row=target_row, column=detected_cols[target_clean]).value = row_pim[pim_col]
                
                # Rellenar Fijos
                for amz_key, val in VALORES_FIJOS.items():
                    target_clean = super_clean(amz_key)
                    if target_clean in detected_cols:
                        ws.cell(row=target_row, column=detected_cols[target_clean]).value = val
                count += 1

            # --- DESCARGA ---
            st.success(f"Se han escrito {count} filas.")
            out = BytesIO()
            wb.save(out)
            st.download_button("📥 Descargar Resultado", out.getvalue(), "Resultado.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")