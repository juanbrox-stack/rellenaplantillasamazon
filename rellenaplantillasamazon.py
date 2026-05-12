import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import json

st.set_page_config(page_title="Amazon Bulk Master", layout="wide")

st.title("🚀 Amazon Template Automator - Versión Completa")

# --- 1. CONFIGURACIÓN DE VALORES FIJOS ---
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

# --- 2. GESTIÓN DEL MAPEO (ESTADO DE SESIÓN) ---
if 'mapeo' not in st.session_state:
    st.session_state.mapeo = {
        "vendor_sku#1.value": "SKU",
        "external_product_id#1.value": "EAN ",
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

# --- 3. BARRA LATERAL ---
st.sidebar.header("⚙️ Configuración de Mapeo")

# Guardar/Cargar
with st.sidebar.expander("💾 Guardar/Cargar Configuración"):
    config_json = json.dumps(st.session_state.mapeo, indent=4)
    st.download_button("📥 Descargar Mapeo Actual", config_json, "mapeo_amazon.json")
    up_config = st.file_uploader("📤 Cargar Mapeo Guardado", type=["json"])
    if up_config:
        st.session_state.mapeo = json.load(up_config)
        st.rerun()

# Añadir nuevo
with st.sidebar.expander("➕ Añadir Nueva Equivalencia"):
    n_amz = st.text_input("Etiqueta Amazon (Fila 4)")
    n_pim = st.text_input("Columna PIM")
    if st.button("Añadir"):
        if n_amz and n_pim:
            st.session_state.mapeo[n_amz] = n_pim
            st.rerun()

# Lista editable (Muestra TODO)
st.sidebar.markdown("---")
st.sidebar.subheader("📝 Mapeo de Columnas")
final_mapping = {}
for amz, pim in st.session_state.mapeo.items():
    final_mapping[amz] = st.sidebar.text_input(f"Amazon: {amz}", value=pim, key=f"key_{amz}")

# --- 4. PROCESAMIENTO ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Procesar Plantilla"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            df_pim.columns = [str(c).strip() for c in df_pim.columns]
            
            wb = load_workbook(amz_template)
            ws = next((wb[n] for n in wb.sheetnames if "template" in n.lower()), wb.worksheets[1])

            # Detectar Columnas (Fila 4)
            amz_cols = {str(ws.cell(row=4, column=c).value).strip(): c 
                        for c in range(1, ws.max_column + 1) if ws.cell(row=4, column=c).value}

            # Escribir (Fila 7)
            for i, row_pim in df_pim.iterrows():
                r_idx = i + 7
                # Variables
                for a_key, p_col in final_mapping.items():
                    p_col = p_col.strip()
                    if a_key in amz_cols and p_col in df_pim.columns:
                        ws.cell(row=r_idx, column=amz_cols[a_key]).value = row_pim[p_col]
                
                # Fijos
                for a_key, val in VALORES_FIJOS.items():
                    if a_key in amz_cols:
                        ws.cell(row=r_idx, column=amz_cols[a_key]).value = val

            output = BytesIO()
            wb.save(output)
            st.success("✅ ¡Procesado con éxito!")
            st.download_button("📥 Descargar Archivo Final", output.getvalue(), "Amazon_Bulk_Fill.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")