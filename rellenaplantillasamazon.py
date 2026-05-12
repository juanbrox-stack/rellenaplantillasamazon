import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Bulk Tool v3", layout="wide")

st.title("🚀 Amazon Template Automator - Corrección EAN y Bullets")

# --- 1. VALORES FIJOS (Campos que Amazon necesita tal cual) ---
VALORES_FIJOS = {
    "external_product_id#1.type": "EAN",  # Esto indica que lo que sigue es un EAN
    "brand#1.value": "Cecotec",
    "package_level#1.value": "Unit",
    "is_trade_item_orderable_unit#1.value": "No",
    "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain",
    "item_weight#1.unit": "Kilograms",
    "item_package_weight#1.unit": "Kilograms",
    "wattage#1.unit": "Watts"
}

# --- 2. MAPEO VARIABLE (Lo que sacamos de tu PIM) ---
# He ajustado "external_product_id#1.value" para que tome el número del PIM
MAPEO_INICIAL = {
    "vendor_sku#1.value": "SKU",
    "external_product_id#1.value": "EAN",  # <--- AQUÍ VA EL NÚMERO (Columna C de tu PIM)
    "merchant_suggested_asin#1.value": "ASIN",
    "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)",
    "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "bullet_point#3.value": "Bulletpoint 3 (FR)",
    "bullet_point#4.value": "Bulletpoint 4 (FR)",
    "bullet_point#5.value": "Bulletpoint 5 (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)",
    "item_package_dimensions#1.length.value": "Largo Caja Color cm",
    "item_package_dimensions#1.width.value": "Ancho Caja Color cm",
    "item_package_dimensions#1.height.value": "Alto Caja Color cm",
    "item_package_weight#1.value": "Peso Caja Color (kg)",
    "item_weight#1.value": "Product weight (Kg)",
    "wattage#1.value": "Potencia (W)"
}

# --- BARRA LATERAL PARA AJUSTES ---
st.sidebar.header("⚙️ Ajustes de Mapeo PIM")
final_mapping = {}
with st.sidebar.expander("📝 Verificar nombres de columnas PIM"):
    for amz, pim in MAPEO_INICIAL.items():
        final_mapping[amz] = st.text_input(f"Amazon: {amz}", value=pim, key=f"map_{amz}")

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM (Excel/CSV)", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon (Excel)", type=["xlsx"])

if st.button("🚀 Generar Fichero Corregido"):
    if pim_file and amz_template:
        try:
            # Leer PIM
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            
            # Cargar Plantilla
            wb = load_workbook(amz_template)
            ws = None
            for name in wb.sheetnames:
                if "template" in name.lower():
                    ws = wb[name]
                    break
            if not ws: ws = wb.worksheets[1]

            # Detectar columnas en la Fila 4
            amazon_cols = {str(ws.cell(row=4, column=c).value).strip(): c 
                           for c in range(1, ws.max_column + 1) if ws.cell(row=4, column=c).value}

            # Procesar filas (desde la 7)
            success_count = 0
            for i, row_pim in df_pim.iterrows():
                row_idx = i + 7
                
                # 1. Rellenar datos del PIM
                for amz_key, pim_col in final_mapping.items():
                    if amz_key in amazon_cols and pim_col in df_pim.columns:
                        ws.cell(row=row_idx, column=amazon_cols[amz_key]).value = row_pim[pim_col]
                
                # 2. Rellenar valores fijos (incluye el tipo "EAN")
                for amz_key, val in VALORES_FIJOS.items():
                    if amz_key in amazon_cols:
                        ws.cell(row=row_idx, column=amazon_cols[amz_key]).value = val
                
                success_count += 1

            # Descarga
            output = BytesIO()
            wb.save(output)
            st.success(f"✅ Procesado: {success_count} productos. EAN y Bullets 3-5 incluidos.")
            st.download_button("📥 Descargar Plantilla Amazon", output.getvalue(), "Amazon_Final_Correcto.xlsx")

        except Exception as e:
            st.error(f"Error técnico: {e}")