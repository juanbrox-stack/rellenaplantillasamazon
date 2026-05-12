import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Template Pro", layout="wide")

st.title("🛠️ Automatizador de Plantillas Amazon (Carga Masiva)")

# --- 1. CONFIGURACIÓN INICIAL (MAPEOS) ---

# Valores que siempre son iguales
VALORES_FIJOS = {
    "brand#1.value": "Cecotec",
    "external_product_id#1.type": "EAN",
    "package_level#1.value": "Unit",
    "is_trade_item_orderable_unit#1.value": "No",
    "manufacturer#1.value": "Cecotec",
    "number_of_items#1.value": 1,
    "is_oem_authorized#1.value": "Yes",
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
    "item_package_weight#1.unit": "Kilograms",
    "rtip_items_per_inner_pack#1.value": 1,
    "country_of_origin#1.value": "Spain",
    "warranty_description#1.value": "2 ans du garantie",
    "supplier_declared_dg_hz_regulation#1.value": "Not Applicable",
    "item_weight#1.unit": "Kilograms",
    "eu_spare_part_availability_duration#1.value": 10,
    "eu_spare_part_availability_duration#1.unit": "Years",
    "dsa_responsible_party_address#1.value": "https://cecotec.es/",
    "compliance_media#1.content_type": "User Manual",
    "compliance_media#1.content_language": "fr_FR",
    "gpsr_safety_attestation#1.value": "Yes",
    "gpsr_manufacturer_reference#1.gpsr_manufacturer_email_address": "https://cecotec.es/"
}

# Mapeo Dinámico: { "Columna Amazon": "Columna PIM" }
MAPEO_VARIABLES = {
    "vendor_sku#1.value": "SKU",
    "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN",
    "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)",
    "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "bullet_point#3.value": "Bulletpoint 3 (FR)",
    "bullet_point#4.value": "Bulletpoint 4 (FR)",
    "bullet_point#5.value": "Bulletpoint 5 (FR)",
    "item_type_name#1.value": "Subfamilia (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)",
    "color#1.value": "color",
    "part_number#1.value": "SKU",
    "oem_equivalent_part_number#1.value": "SKU",
    "wattage#1.value": "Potencia (W)",
    "included_components#1.value": "Contenido de la caja (FR)",
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

# --- 2. INTERFAZ ---

st.sidebar.header("🔧 Ajustes de Mapeo")
with st.sidebar.expander("Editar Mapeo PIM -> Amazon"):
    edited_mapping = {}
    for amz, pim in MAPEO_VARIABLES.items():
        new_pim = st.text_input(f"Amazon: {amz}", value=pim)
        edited_mapping[amz] = new_pim

col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 Subir Fichero PIM (FRIGOS.xlsx)", type=["xlsx"])
with col2:
    amz_template = st.file_uploader("📂 Subir Plantilla Amazon Original", type=["xlsx"])

if st.button("🚀 Procesar y Cumplimentar"):
    if pim_file and amz_template:
        try:
            # Leer PIM
            df_pim = pd.read_excel(pim_file)
            
            # Cargar Plantilla Amazon con openpyxl (mantiene formato)
            wb = load_workbook(amz_template)
            # La pestaña de carga suele ser la segunda (índice 1)
            ws = wb.worksheets[1] 
            
            # Obtener las cabeceras técnicas de la fila 3 de la plantilla
            amazon_headers = [cell.value for cell in ws[3]]
            
            # Diccionario para saber en qué columna (letra/índice) está cada campo de Amazon
            col_idx = {header: i+1 for i, header in enumerate(amazon_headers) if header}

            # --- PROCESO DE ESCRITURA ---
            for i, row_pim in df_pim.iterrows():
                excel_row = i + 4 # Empezamos en la fila 4 (debajo de cabeceras)
                
                # A. Escribir Valores Fijos
                for amz_col, val in VALORES_FIJOS.items():
                    if amz_col in col_idx:
                        ws.cell(row=excel_row, column=col_idx[amz_col]).value = val
                
                # B. Escribir Valores Variables (PIM)
                for amz_col, pim_col in edited_mapping.items():
                    if amz_col in col_idx and pim_col in df_pim.columns:
                        ws.cell(row=excel_row, column=col_idx[amz_col]).value = row_pim[pim_col]

            # Guardar resultado
            out_buf = BytesIO()
            wb.save(out_buf)
            
            st.success(f"✅ Se han rellenado {len(df_pim)} filas en la plantilla original.")
            st.download_button(
                label="📥 Descargar Plantilla Cumplimentada",
                data=out_buf.getvalue(),
                file_name="Amazon_Template_FULL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")