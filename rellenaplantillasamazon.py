import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Bulk Tool", layout="wide")

st.title("🛠️ Automatizador de Plantillas Amazon")
st.markdown("Carga cualquier plantilla de Amazon y tus datos de PIM para generar el fichero de carga masiva.")

# --- CONFIGURACIÓN DE VALORES FIJOS (Por defecto) ---
VALORES_FIJOS_DEFAULT = {
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

# --- CONFIGURACIÓN DE MAPEO VARIABLE (Por defecto) ---
MAPEO_VARIABLES_DEFAULT = {
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

# --- BARRA LATERAL: EDICIÓN DE MAPEO ---
st.sidebar.header("⚙️ Configuración de Mapeo")
st.sidebar.info("Modifica los nombres de las columnas del PIM si no coinciden con la categoría actual.")

with st.sidebar.expander("📝 Editar Correspondencias"):
    current_mapping = {}
    for amz, pim_default in MAPEO_VARIABLES_DEFAULT.items():
        val = st.text_input(f"Amazon: {amz}", value=pim_default, key=amz)
        current_mapping[amz] = val

# --- ÁREA DE CARGA ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Fichero Origen (PIM)", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Destino (Amazon)", type=["xlsx"])

if st.button("🚀 Generar Fichero de Carga Masiva"):
    if pim_file and amz_template:
        try:
            # Leer PIM
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            
            # Cargar Plantilla Amazon (Formato original)
            wb = load_workbook(amz_template)
            # Buscamos la pestaña de datos (Template)
            sheet_name = "Template" if "Template" in wb.sheetnames else wb.sheetnames[1]
            ws = wb[sheet_name]
            
            # Detectar Columnas en Fila 4 (Nombres técnicos)
            amazon_cols = {str(cell.value).strip(): i+1 for i, cell in enumerate(ws[4]) if cell.value}
            
            if not amazon_cols:
                st.error("No se han encontrado etiquetas técnicas en la fila 4 de la plantilla.")
            else:
                rows_added = 0
                # Procesar cada fila del PIM
                for i, row_pim in df_pim.iterrows():
                    target_row = i + 7 # Datos empiezan en fila 7
                    
                    # Rellenar Variables mapeadas
                    for amz_key, pim_col in current_mapping.items():
                        if amz_key in amazon_cols and pim_col in df_pim.columns:
                            ws.cell(row=target_row, column=amazon_cols[amz_key]).value = row_pim[pim_col]
                    
                    # Rellenar Valores Fijos
                    for amz_key, fixed_val in VALORES_FIJOS_DEFAULT.items():
                        if amz_key in amazon_cols:
                            ws.cell(row=target_row, column=amazon_cols[amz_key]).value = fixed_val
                    
                    rows_added += 1

                # Preparar descarga
                output = BytesIO()
                wb.save(output)
                
                st.success(f"✅ ¡Completado! Se han rellenado {rows_added} filas.")
                st.download_button(
                    label="📥 Descargar Plantilla Cumplimentada",
                    data=output.getvalue(),
                    file_name="Carga_Masiva_Amazon_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Se produjo un error: {e}")
    else:
        st.warning("Por favor, sube ambos archivos para procesar.")