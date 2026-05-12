import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Pro Filler", layout="wide")

st.title("🚀 Amazon Template Automator")

# --- DICCIONARIOS POR DEFECTO ---
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
    "item_package_dimensions#1.width.value": "Ancho Caja Color cm",
    "item_package_dimensions#1.height.value": "Alto Caja Color cm",
    "item_package_weight#1.value": "Peso Caja Color (kg)",
    "item_weight#1.value": "Product weight (Kg)",
    "compliance_media#1.source_location": "Manual de instrucciones (Comercial)"
}

# --- INTERFAZ ---
st.sidebar.header("⚙️ Configuración")
with st.sidebar.expander("📝 Editar Mapeo PIM"):
    final_mapping = {}
    for amz, pim in MAPEO_VARIABLES.items():
        res = st.text_input(f"AMZ: {amz}", value=pim)
        final_mapping[amz] = res

col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Ejecutar y Rellenar"):
    if pim_file and amz_template:
        try:
            # Leer PIM
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            
            # Cargar Plantilla original con openpyxl
            wb = load_workbook(amz_template)
            # Buscar la pestaña de carga (segunda o por nombre)
            sheet_name = "Template" if "Template" in wb.sheetnames else wb.sheetnames[1]
            ws = wb[sheet_name]
            
            # --- DIAGNÓSTICO DE COLUMNAS ---
            # Leemos la fila 4 para ver los nombres técnicos
            amazon_cols = {}
            raw_headers = []
            for i, cell in enumerate(ws[4], 1):
                if cell.value:
                    header_name = str(cell.value).strip()
                    amazon_cols[header_name] = i
                    raw_headers.append(header_name)
            
            st.write("### 🔍 Diagnóstico de Plantilla")
            st.write(f"Columnas técnicas encontradas en Fila 4: `{len(raw_headers)}`")
            
            # Verificar emparejamiento
            missing = [k for k in final_mapping.keys() if k not in amazon_cols]
            if missing:
                with st.expander("⚠️ Columnas de Amazon NO encontradas en tu plantilla"):
                    st.write(missing)
            
            # --- ESCRITURA ---
            rows_filled = 0
            for i, row_pim in df_pim.iterrows():
                current_row = i + 7 # Amazon suele empezar datos en fila 7
                
                # Escribir mapeo dinámico
                for amz_key, pim_col in final_mapping.items():
                    if amz_key in amazon_cols and pim_col in df_pim.columns:
                        ws.cell(row=current_row, column=amazon_cols[amz_key]).value = row_pim[pim_col]
                
                # Escribir fijos
                for amz_key, val in VALORES_FIJOS.items():
                    if amz_key in amazon_cols:
                        ws.cell(row=current_row, column=amazon_cols[amz_key]).value = val
                
                rows_filled += 1

            # --- DESCARGA ---
            output = BytesIO()
            wb.save(output)
            st.success(f"✅ Se han rellenado {rows_filled} filas.")
            st.download_button("📥 Descargar Archivo Final", output.getvalue(), "Amazon_Relleno.xlsx")

        except Exception as e:
            st.error(f"Error técnico: {e}")