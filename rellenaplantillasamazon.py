import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Amazon Bulk Tool", layout="wide")

st.title("🚀 Amazon Template Automator")
st.success("Estructura detectada: Atributos en Fila 4 | Datos en Fila 7")

# --- DICCIONARIOS DE CONFIGURACIÓN ---
VALORES_FIJOS = {
    "brand#1.value": "Cecotec", "external_product_id#1.type": "EAN", 
    "package_level#1.value": "Unit", "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain", "item_weight#1.unit": "Kilograms",
    "wattage#1.unit": "Watts", "item_package_weight#1.unit": "Kilograms"
}

MAPEO_VARIABLES = {
    "vendor_sku#1.value": "SKU", "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN", "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)", "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Ajustes de Mapeo PIM")
user_mapping = {}
with st.sidebar.expander("📝 Editar Correspondencias"):
    for amz, pim in MAPEO_VARIABLES.items():
        user_mapping[amz] = st.text_input(f"Amazon: {amz}", value=pim, key=amz)

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Generar Fichero"):
    if pim_file and amz_template:
        try:
            # 1. Leer PIM
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            
            # 2. Abrir Plantilla Amazon
            wb = load_workbook(amz_template)
            # Seleccionamos la pestaña 'Template' o la segunda
            ws = wb["Template"] if "Template" in wb.sheetnames else wb.worksheets[1]

            # 3. Mapear columnas de la FILA 4 (Atributos Técnicos)
            # Limpiamos espacios y saltos de línea
            amz_col_map = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=4, column=c).value
                if val:
                    amz_col_map[str(val).strip()] = c

            # 4. Proceso de Escritura (Empieza en FILA 7)
            rows_added = 0
            for i, row_pim in df_pim.iterrows():
                current_row = i + 7 
                
                # Rellenar Variables del PIM
                for amz_key, pim_col in user_mapping.items():
                    if amz_key in amz_col_map and pim_col in df_pim.columns:
                        ws.cell(row=current_row, column=amz_col_map[amz_key]).value = row_pim[pim_col]
                
                # Rellenar Valores Fijos
                for amz_key, fixed_val in VALORES_FIJOS.items():
                    if amz_key in amz_col_map:
                        ws.cell(row=current_row, column=amz_col_map[amz_key]).value = fixed_val
                
                rows_added += 1

            # 5. Generar Descarga
            output = BytesIO()
            wb.save(output)
            st.success(f"✅ Se han rellenado {rows_added} filas de productos.")
            st.download_button("📥 Descargar Plantilla Rellena", output.getvalue(), "Amazon_Bulk_Ready.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Sube ambos archivos para procesar.")