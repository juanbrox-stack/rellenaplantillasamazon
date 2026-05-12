import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Full Mapping Tool", layout="wide")

st.title("🚀 Amazon Template Automator - Mapeo Completo")
st.markdown("Verifica en la barra lateral que todos los nombres de columna coincidan con tu Excel PIM.")

# --- 1. VALORES FIJOS (Se mantienen automáticos) ---
VALORES_FIJOS = {
    "external_product_id#1.type": "EAN",
    "brand#1.value": "Cecotec",
    "package_level#1.value": "Unit",
    "is_trade_item_orderable_unit#1.value": "No",
    "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain",
    "item_weight#1.unit": "Kilograms",
    "item_package_weight#1.unit": "Kilograms",
    "wattage#1.unit": "Watts",
    "item_depth_width_height#1.depth.unit": "Centimeters",
    "item_depth_width_height#1.height.unit": "Centimeters",
    "item_depth_width_height#1.width.unit": "Centimeters",
    "item_dimensions#1.length.unit": "Centimeters",
    "item_dimensions#1.width.unit": "Centimeters",
    "item_dimensions#1.height.unit": "Centimeters",
    "item_package_dimensions#1.length.unit": "Centimeters",
    "item_package_dimensions#1.width.unit": "Centimeters",
    "item_package_dimensions#1.height.unit": "Centimeters",
    "eu_spare_part_availability_duration#1.value": 10,
    "eu_spare_part_availability_duration#1.unit": "Years",
    "dsa_responsible_party_address#1.value": "https://cecotec.es/",
    "compliance_media#1.content_type": "User Manual",
    "compliance_media#1.content_language": "fr_FR",
    "gpsr_safety_attestation#1.value": "Yes",
    "gpsr_manufacturer_reference#1.gpsr_manufacturer_email_address": "https://cecotec.es/"
}

# --- 2. MAPEO VARIABLE (Ahora con TODOS los campos solicitados) ---
MAPEO_INICIAL = {
    "vendor_sku#1.value": "SKU",
    "external_product_id#1.value": "EAN ", # Nota: Tu Excel tiene un espacio después de EAN
    "merchant_suggested_asin#1.value": "ASIN PIM",
    "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)",
    "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "bullet_point#3.value": "Bulletpoint 3 (FR)",
    "bullet_point#4.value": "Bulletpoint 4 (FR)",
    "bullet_point#5.value": "Bulletpoint 5 (FR)",
    "item_type_name#1.value": "Subfamilia (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)",
    "color#1.value": "Color",
    "part_number#1.value": "SKU",
    "oem_equivalent_part_number#1.value": "SKU",
    "wattage#1.value": "Potencia (W)",
    "included_components#1.value": "Contenido de la caja",
    "item_depth_width_height#1.depth.value": "Product depth (cm)",
    "item_depth_width_height#1.height.value": "Product height (cm)",
    "item_depth_width_height#1.width.value": "Product width (cm)",
    "website_shipping_weight#1.value": "Peso Caja Color (kg)",
    "item_dimensions#1.length.value": "Product depth (cm)",
    "item_dimensions#1.width.value": "Product width (cm)",
    "item_dimensions#1.height.value": "Product height (cm)",
    "item_package_dimensions#1.length.value": "Largo Caja Color cm",
    "item_package_dimensions#1.width.value": "Ancho Caja Color cm",
    "item_package_dimensions#1.height.value": "Alto Caja Color cm",
    "item_package_weight#1.value": "Peso Caja Color (kg)",
    "item_weight#1.value": "Product weight (Kg)",
    "compliance_media#1.source_location": "Manual de instrucciones (Comercial)"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Ajustes de Mapeo PIM")
final_mapping = {}
with st.sidebar.expander("📝 Verificar nombres de columnas PIM", expanded=True):
    for amz, pim in MAPEO_INICIAL.items():
        # Creamos un input para cada campo, ahora saldrán todos
        final_mapping[amz] = st.text_input(f"Amazon: {amz}", value=pim, key=f"f_{amz}")

# --- CARGA ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Generar Plantilla Final"):
    if pim_file and amz_template:
        try:
            # Leer PIM (Detectando si es Excel o CSV)
            if pim_file.name.endswith('.csv'):
                df_pim = pd.read_csv(pim_file)
            else:
                df_pim = pd.read_excel(pim_file)
            
            # Limpiar espacios en los nombres de las columnas del PIM para evitar errores
            df_pim.columns = [str(c).strip() for c in df_pim.columns]

            wb = load_workbook(amz_template)
            ws = None
            for name in wb.sheetnames:
                if "template" in name.lower():
                    ws = wb[name]
                    break
            if not ws: ws = wb.worksheets[1]

            # Mapear Fila 4 de la plantilla de Amazon
            amazon_cols = {}
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=4, column=c).value
                if val:
                    amazon_cols[str(val).strip()] = c

            # Escritura en Fila 7
            rows_filled = 0
            for i, row_pim in df_pim.iterrows():
                row_idx = i + 7
                
                # 1. Variables (Mapeo de la barra lateral)
                for amz_key, pim_col in final_mapping.items():
                    pim_col_clean = str(pim_col).strip()
                    if amz_key in amazon_cols and pim_col_clean in df_pim.columns:
                        ws.cell(row=row_idx, column=amazon_cols[amz_key]).value = row_pim[pim_col_clean]
                
                # 2. Valores Fijos
                for amz_key, val in VALORES_FIJOS.items():
                    if amz_key in amazon_cols:
                        ws.cell(row=row_idx, column=amazon_cols[amz_key]).value = val
                
                rows_filled += 1

            # Generar archivo
            output = BytesIO()
            wb.save(output)
            st.success(f"✅ ¡Todo listo! Se han procesado {rows_filled} productos con el mapeo completo.")
            st.download_button("📥 Descargar Amazon_Final_Completo.xlsx", output.getvalue(), "Amazon_Final_Completo.xlsx")

        except Exception as e:
            st.error(f"Error técnico: {e}")