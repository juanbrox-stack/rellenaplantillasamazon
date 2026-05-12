import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Template Pro", layout="wide")

st.title("📦 Amazon Template Automator")
st.info("Configurado para leer nombres técnicos en Fila 4 y escribir datos desde Fila 7.")

# --- DICCIONARIOS (Tus datos) ---
VALORES_FIJOS = {
    "brand#1.value": "Cecotec", "external_product_id#1.type": "EAN", 
    "package_level#1.value": "Unit", "is_trade_item_orderable_unit#1.value": "No",
    "manufacturer#1.value": "Cecotec", "number_of_items#1.value": 1,
    "is_oem_authorized#1.value": "Yes", "wattage#1.unit": "Watts",
    "item_depth_width_height#1.depth.unit": "Centimeters",
    "item_depth_width_height#1.height.unit": "Centimeters",
    "item_depth_width_height#1.width.unit": "Centimeters",
    "item_dimensions#1.length.unit": "Centimeters",
    "item_dimensions#1.width.unit": "Centimeters",
    "item_dimensions#1.height.unit": "Centimeters",
    "item_package_dimensions#1.length.unit": "Centimeters",
    "item_package_dimensions#1.width.unit": "Centimeters",
    "item_package_dimensions#1.height.unit": "Centimeters",
    "item_package_weight#1.unit": "Kilograms", "rtip_items_per_inner_pack#1.value": 1,
    "country_of_origin#1.value": "Spain", "warranty_description#1.value": "2 ans du garantie",
    "supplier_declared_dg_hz_regulation#1.value": "Not Applicable",
    "item_weight#1.unit": "Kilograms", "eu_spare_part_availability_duration#1.value": 10,
    "eu_spare_part_availability_duration#1.unit": "Years",
    "dsa_responsible_party_address#1.value": "https://cecotec.es/",
    "compliance_media#1.content_type": "User Manual",
    "compliance_media#1.content_language": "fr_FR", "gpsr_safety_attestation#1.value": "Yes",
    "gpsr_manufacturer_reference#1.gpsr_manufacturer_email_address": "https://cecotec.es/"
}

MAPEO_VARIABLES = {
    "vendor_sku#1.value": "SKU", "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN", "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)", "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "bullet_point#3.value": "Bulletpoint 3 (FR)", "bullet_point#4.value": "Bulletpoint 4 (FR)",
    "bullet_point#5.value": "Bulletpoint 5 (FR)", "item_type_name#1.value": "Subfamilia (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)",
    "color#1.value": "color", "part_number#1.value": "SKU",
    "oem_equivalent_part_number#1.value": "SKU", "wattage#1.value": "Potencia (W)",
    "included_components#1.value": "Contenido de la caja (FR)",
    "item_depth_width_height#1.depth.value": "Product depth (cm)",
    "item_depth_width_height#1.height.value": "Product height (cm)",
    "item_depth_width_height#1.width.value": "Product width (cm)",
    "website_shipping_weight#1.value": "Peso Caja Color (kg)",
    "item_dimensions#1.length.value": "Product depth (cm)",
    "item_dimensions#1.width.value": "Product width (cm)",
    "item_dimensions#1.height.value": "Product height (cm)",
    "item_package_dimensions#1.length.value": "Largo Caja Color cm",
    "item_package_dimensions#1.width.unit": "Ancho Caja Color cm", # Corregido de .unit a .value si es dato
    "item_package_dimensions#1.height.value": "Alto Caja Color cm",
    "item_package_weight#1.value": "Peso Caja Color (kg)",
    "item_weight#1.value": "Product weight (Kg)",
    "compliance_media#1.source_location": "Manual de instrucciones (Comercial)"
}

col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 Subir PIM (FRIGOS)", type=["xlsx"])
with col2:
    amz_template = st.file_uploader("📂 Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Ejecutar"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file)
            wb = load_workbook(amz_template)
            # Buscamos la pestaña 'Template' por posición o nombre
            ws = wb["Template"] if "Template" in wb.sheetnames else wb.worksheets[1]
            
            # --- DETECCIÓN DE COLUMNAS (FILA 4) ---
            # Guardamos la posición de cada columna técnica
            amazon_cols = {str(cell.value).strip(): i+1 for i, cell in enumerate(ws[4]) if cell.value}
            
            if not amazon_cols:
                st.error("❌ No se detectaron nombres técnicos en la fila 4. Verifica el archivo.")
            else:
                count_mapped = 0
                # --- ESCRITURA DE DATOS (EMPIEZA FILA 7) ---
                for i, row_pim in df_pim.iterrows():
                    current_row = i + 7 
                    
                    # 1. Variables PIM
                    for amz_key, pim_col in MAPEO_VARIABLES.items():
                        if amz_key in amazon_cols and pim_col in df_pim.columns:
                            ws.cell(row=current_row, column=amazon_cols[amz_key]).value = row_pim[pim_col]
                            if i == 0: count_mapped += 1
                    
                    # 2. Valores Fijos
                    for amz_key, val in VALORES_FIJOS.items():
                        if amz_key in amazon_cols:
                            ws.cell(row=current_row, column=amazon_cols[amz_key]).value = val

                st.success(f"✅ Proceso finalizado. Mapeadas {count_mapped} columnas.")
                
                # --- DESCARGA ---
                out = BytesIO()
                wb.save(out)
                st.download_button("📥 Descargar Archivo", out.getvalue(), "Amazon_Final_Relleno.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")