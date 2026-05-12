import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import json

st.set_page_config(page_title="Amazon Custom Mapper", layout="wide")

st.title("🚀 Amazon Template Automator - Mapeo Evolutivo")

# --- 1. GESTIÓN DE ESTADO (MEMORIA DE LA APP) ---
# Inicializamos el mapeo en la sesión de Streamlit si no existe
if 'mapeo_personalizado' not in st.session_state:
    st.session_state.mapeo_personalizado = {
        "vendor_sku#1.value": "SKU",
        "external_product_id#1.value": "EAN ",
        "merchant_suggested_asin#1.value": "ASIN PIM",
        "model_number#1.value": "Nombre Producto / Modelo",
        "bullet_point#1.value": "Bulletpoint 1 (FR)",
        "bullet_point#2.value": "Bulletpoint 2 (FR)",
        "bullet_point#3.value": "Bulletpoint 3 (FR)",
        "bullet_point#4.value": "Bulletpoint 4 (FR)",
        "bullet_point#5.value": "Bulletpoint 5 (FR)",
        "rtip_product_description#1.value": "Descripción larga del producto (FR)",
        "wattage#1.value": "Potencia (W)"
    }

# --- 2. BARRA LATERAL: CONFIGURACIÓN Y PERSISTENCIA ---
st.sidebar.header("⚙️ Gestión de Mapeo")

# Opción para guardar/cargar la configuración en un archivo
with st.sidebar.expander("💾 Guardar/Cargar Configuración"):
    # Botón para descargar el mapeo actual como JSON
    mapeo_json = json.dumps(st.session_state.mapeo_personalizado, indent=4)
    st.download_button("📥 Descargar mi Mapeo (JSON)", mapeo_json, "mi_mapeo_amazon.json")
    
    # Opción para subir un mapeo guardado previamente
    uploaded_config = st.file_uploader("📤 Cargar Mapeo previo", type=["json"])
    if uploaded_config:
        st.session_state.mapeo_personalizado = json.load(uploaded_config)
        st.success("Configuración cargada.")

# Formulario para AÑADIR NUEVAS EQUIVALENCIAS
with st.sidebar.expander("➕ Añadir Nueva Equivalencia"):
    new_amz = st.text_input("Nombre técnico Amazon (ej: color#1.value)")
    new_pim = st.text_input("Nombre columna PIM (ej: Color)")
    if st.button("Añadir al mapeo"):
        if new_amz and new_pim:
            st.session_state.mapeo_personalizado[new_amz] = new_pim
            st.rerun()

# Listado editable de mapeos actuales
st.sidebar.markdown("---")
st.sidebar.subheader("📝 Mapeo Actual")
final_mapping = {}
for amz, pim in st.session_state.mapeo_personalizado.items():
    final_mapping[amz] = st.sidebar.text_input(f"AMZ: {amz}", value=pim, key=f"inp_{amz}")

# --- 3. PROCESAMIENTO ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Generar Plantilla con Mapeo Personalizado"):
    if pim_file and amz_template:
        try:
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            df_pim.columns = [str(c).strip() for c in df_pim.columns]
            
            wb = load_workbook(amz_template)
            ws = next((wb[n] for n in wb.sheetnames if "template" in n.lower()), wb.worksheets[1])

            # Detectar columnas en Fila 4
            amazon_cols = {str(ws.cell(row=4, column=c).value).strip(): c 
                           for c in range(1, ws.max_column + 1) if ws.cell(row=4, column=c).value}

            # Escritura (Fila 7)
            for i, row_pim in df_pim.iterrows():
                row_idx = i + 7
                for amz_key, pim_col in final_mapping.items():
                    if amz_key in amazon_cols and pim_col in df_pim.columns:
                        ws.cell(row=row_idx, column=amazon_cols[amz_key]).value = row_pim[pim_col]
                
                # Rellenar también los VALORES FIJOS (puedes definirlos aquí o en otro dict)
                ws.cell(row=row_idx, column=amazon_cols.get("external_product_id#1.type", 0)).value = "EAN"
                # ... añadir aquí el resto de fijos si no cambian nunca

            output = BytesIO()
            wb.save(output)
            st.success("✅ ¡Procesado con éxito!")
            st.download_button("📥 Descargar Excel", output.getvalue(), "Amazon_Custom.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")