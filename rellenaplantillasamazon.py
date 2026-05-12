import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Amazon Auto-Mapper", layout="wide")
st.title("🚀 Amazon Template Automator")

# --- CONFIGURACIÓN DE MAPEOS (Se mantiene de tus mensajes previos) ---
VALORES_FIJOS = {
    "brand#1.value": "Cecotec", "external_product_id#1.type": "EAN", 
    "package_level#1.value": "Unit", "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain", "item_weight#1.unit": "Kilograms"
}

MAPEO_INICIAL = {
    "vendor_sku#1.value": "SKU", "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN", "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)", "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Configuración de Mapeo")
final_mapping = {}
with st.sidebar.expander("📝 Editar Correspondencias PIM"):
    for amz, pim in MAPEO_INICIAL.items():
        final_mapping[amz] = st.text_input(f"Amazon: {amz}", value=pim)

# --- FUNCIONES ---
def super_clean(text):
    if text is None: return ""
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

# --- CARGA ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Ejecutar Proceso"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            wb = load_workbook(amz_template)
            sheet_name = "Template" if "Template" in wb.sheetnames else wb.sheetnames[1]
            ws = wb[sheet_name]
            
            # --- MOTOR DE AUTODETECCIÓN DE FILA TÉCNICA ---
            detected_cols = {}
            found_row = None
            
            # Buscamos en las filas 3, 4 y 5 la que contenga "sku" o "product"
            for row_num in [3, 4, 5]:
                temp_cols = {super_clean(ws.cell(row=row_num, column=c).value): c 
                             for c in range(1, ws.max_column + 1) if ws.cell(row=row_num, column=c).value}
                if any("sku" in k or "productid" in k for k in temp_cols.keys()):
                    detected_cols = temp_cols
                    found_row = row_num
                    break
            
            if not found_row:
                st.error("❌ No se pudo detectar la fila técnica de Amazon (Fila 3, 4 o 5).")
            else:
                st.info(f"🔍 Detectados nombres técnicos en la **Fila {found_row}**.")
                
                # --- ESCRITURA ---
                rows_filled = 0
                for i, row_pim in df_pim.iterrows():
                    target_row = i + 7 # Amazon siempre empieza datos en 7
                    
                    # 1. Variables (Barra Lateral)
                    for amz_key, pim_col in final_mapping.items():
                        clean_key = super_clean(amz_key)
                        if clean_key in detected_cols and pim_col in df_pim.columns:
                            ws.cell(row=target_row, column=detected_cols[clean_key]).value = row_pim[pim_col]
                    
                    # 2. Valores Fijos
                    for amz_key, val in VALORES_FIJOS.items():
                        clean_key = super_clean(amz_key)
                        if clean_key in detected_cols:
                            ws.cell(row=target_row, column=detected_cols[clean_key]).value = val
                    rows_filled += 1

                # --- DESCARGA ---
                output = BytesIO()
                wb.save(output)
                st.success(f"✅ Se han rellenado {rows_filled} filas correctamente.")
                st.download_button("📥 Descargar Plantilla Rellena", output.getvalue(), "Amazon_Final_Relleno.xlsx")

        except Exception as e:
            st.error(f"Error técnico: {e}")