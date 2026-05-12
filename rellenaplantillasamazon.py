import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import re

st.set_page_config(page_title="Amazon Bulk Tool", layout="wide")

# --- FUNCIONES DE APOYO ---
def normalize(text):
    """Limpia etiquetas para que 'vendor_sku#1.value' sea igual a 'vendor_sku'"""
    if not text: return ""
    text = str(text).lower().strip()
    text = re.sub(r'#\d+', '', text) # Quita #1, #2...
    text = re.sub(r'\.value|\.unit|\.type', '', text) # Quita extensiones técnicas
    return re.sub(r'[^a-z0-9]', '', text) # Deja solo letras y números

# --- CONFIGURACIÓN POR DEFECTO ---
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
    "item_type_name#1.value": "Subfamilia (FR)", "color#1.value": "color"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Configuración de Mapeo")
final_mapping = {}
with st.sidebar.expander("📝 Editar Correspondencias PIM"):
    for amz, pim in MAPEO_VARIABLES.items():
        final_mapping[amz] = st.text_input(f"AMZ: {amz}", value=pim, key=f"sidebar_{amz}")

# --- INTERFAZ PRINCIPAL ---
st.title("📦 Amazon Template Automator")
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Procesar Plantilla"):
    if pim_file and amz_template:
        try:
            # 1. Cargar datos
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            wb = load_workbook(amz_template)
            # Buscamos la pestaña 'Template'
            ws = wb["Template"] if "Template" in wb.sheetnames else wb.worksheets[1]

            # 2. Mapear Fila 4 (Amazon) de forma flexible
            # Guardamos { 'nombre_normalizado': indice_columna }
            amazon_cols = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=4, column=c).value
                if val:
                    amazon_cols[normalize(val)] = c

            if not amazon_cols:
                st.error("No se detectaron cabeceras en la Fila 4. Intenta con la Fila 3 o revisa el archivo.")
            else:
                # 3. Escritura de Datos
                rows_processed = 0
                log_confirmacion = []

                for i, row_pim in df_pim.iterrows():
                    target_row = i + 7 # Estándar de Amazon: Datos empiezan en 7
                    
                    # Rellenar Variables (Barra Lateral)
                    for amz_tag, pim_col in final_mapping.items():
                        clean_tag = normalize(amz_tag)
                        if clean_tag in amazon_cols and pim_col in df_pim.columns:
                            ws.cell(row=target_row, column=amazon_cols[clean_tag]).value = row_pim[pim_col]
                            if i == 0: log_confirmacion.append(f"Mapeado: {amz_tag} -> Col {amazon_cols[clean_tag]}")

                    # Rellenar Fijos
                    for amz_tag, val_fijo in VALORES_FIJOS.items():
                        clean_tag = normalize(amz_tag)
                        if clean_tag in amazon_cols:
                            ws.cell(row=target_row, column=amazon_cols[clean_tag]).value = val_fijo

                    rows_processed += 1

                # 4. Finalizar
                st.success(f"✅ ¡Hecho! Se han rellenado {rows_processed} filas.")
                with st.expander("Ver detalles del mapeo técnico"):
                    st.write(log_confirmacion)

                output = BytesIO()
                wb.save(output)
                st.download_button("📥 Descargar Resultado", output.getvalue(), "Amazon_Final_Relleno.xlsx")

        except Exception as e:
            st.error(f"Error técnico: {e}")