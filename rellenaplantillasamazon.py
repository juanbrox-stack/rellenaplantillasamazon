import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Amazon Pro Filler", layout="wide")
st.title("🚀 Amazon Template Automator")
st.info("Sistema optimizado para manejar celdas combinadas y filas ocultas.")

# --- CONFIGURACIÓN DE MAPEOS ---
VALORES_FIJOS = {
    "brand#1.value": "Cecotec", "external_product_id#1.type": "EAN", 
    "package_level#1.value": "Unit", "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain", "item_weight#1.unit": "Kilograms",
    "wattage#1.unit": "Watts"
}

MAPEO_VARIABLES = {
    "vendor_sku#1.value": "SKU", "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN", "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)", "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "item_type_name#1.value": "Subfamilia (FR)"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Ajustes de Mapeo PIM")
final_mapping = {}
with st.sidebar.expander("📝 Editar Correspondencias"):
    for amz, pim in MAPEO_VARIABLES.items():
        final_mapping[amz] = st.text_input(f"AMZ: {amz}", value=pim, key=f"v_{amz}")

# --- LIMPIEZA TÉCNICA ---
def super_clean(text):
    if text is None: return ""
    # Quitamos todo lo que no sea letra o número
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

# --- CARGA ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Procesar y Rellenar"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            wb = load_workbook(amz_template, data_only=True) # IMPORTANTE: Leer solo valores
            ws = wb["Template"] if "Template" in wb.sheetnames else wb.worksheets[1]
            
            # --- LOCALIZADOR DE COLUMNAS (FILAS 3 a 5) ---
            amazon_cols = {}
            target_row_tech = None
            
            for row_num in [3, 4, 5]:
                potential_cols = {}
                for c in range(1, ws.max_column + 1):
                    val = ws.cell(row=row_num, column=c).value
                    if val:
                        potential_cols[super_clean(val)] = c
                
                # Si encontramos 'sku' o 'brand', es la fila correcta
                if "vendorsku" in potential_cols or "brand" in potential_cols or "sku" in potential_cols:
                    amazon_cols = potential_cols
                    target_row_tech = row_num
                    break

            if not target_row_tech:
                st.error("❌ No se pudo identificar la fila de etiquetas técnicas (3, 4 o 5).")
            else:
                st.success(f"🔍 Etiquetas detectadas en Fila {target_row_tech}")
                
                # --- ESCRITURA ---
                rows_written = 0
                for i, row_pim in df_pim.iterrows():
                    current_row = i + 7 # Datos empiezan siempre en 7
                    
                    # 1. Variables
                    for amz_key, pim_col in final_mapping.items():
                        clean_k = super_clean(amz_key)
                        if clean_k in amazon_cols and pim_col in df_pim.columns:
                            ws.cell(row=current_row, column=amazon_cols[clean_k]).value = row_pim[pim_col]
                    
                    # 2. Fijos
                    for amz_key, val in VALORES_FIJOS.items():
                        clean_k = super_clean(amz_key)
                        if clean_k in amazon_cols:
                            ws.cell(row=current_row, column=amazon_cols[clean_k]).value = val
                    
                    rows_written += 1

                # --- DESCARGA ---
                output = BytesIO()
                wb.save(output)
                st.success(f"✅ Se han rellenado {rows_written} filas.")
                st.download_button("📥 Descargar Archivo Final", output.getvalue(), "Amazon_Final.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")