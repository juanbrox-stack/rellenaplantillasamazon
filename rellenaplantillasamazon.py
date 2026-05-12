import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Amazon Precision Filler", layout="wide")

st.title("🚀 Amazon Template Automator")
st.info("Estructura: Atributos en Fila 4 | Datos en Fila 7")

# --- CONFIGURACIÓN DE MAPEO (Variables del usuario) ---
VALORES_FIJOS = {
    "brand#1.value": "Cecotec", "external_product_id#1.type": "EAN", 
    "package_level#1.value": "Unit", "manufacturer#1.value": "Cecotec",
    "country_of_origin#1.value": "Spain", "item_weight#1.unit": "Kilograms",
    "wattage#1.unit": "Watts", "item_package_weight#1.unit": "Kilograms",
    "is_trade_item_orderable_unit#1.value": "No"
}

MAPEO_VARIABLES = {
    "vendor_sku#1.value": "SKU", "external_product_id#1.value": "EAN",
    "merchant_suggested_asin#1.value": "ASIN", "model_number#1.value": "Nombre Producto / Modelo",
    "bullet_point#1.value": "Bulletpoint 1 (FR)", "bullet_point#2.value": "Bulletpoint 2 (FR)",
    "rtip_product_description#1.value": "Descripción larga del producto (FR)"
}

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Ajustes de Mapeo PIM")
final_mapping = {}
with st.sidebar.expander("📝 Editar Correspondencias"):
    for amz, pim in MAPEO_VARIABLES.items():
        final_mapping[amz] = st.text_input(f"AMZ: {amz}", value=pim, key=f"p_{amz}")

# --- CARGA ---
col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("📂 1. Subir Datos PIM", type=["xlsx", "csv"])
with col2:
    amz_template = st.file_uploader("📂 2. Subir Plantilla Amazon", type=["xlsx"])

if st.button("🚀 Procesar y Descargar"):
    if pim_file and amz_template:
        try:
            # 1. Leer PIM
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)
            
            # 2. Cargar Plantilla (Carga pesada para asegurar compatibilidad)
            wb = load_workbook(amz_template)
            # Intentar buscar por nombre de pestaña 'Template' o por posición 2 (índice 1)
            sheet_found = False
            for sheet in wb.sheetnames:
                if "template" in sheet.lower():
                    ws = wb[sheet]
                    sheet_found = True
                    break
            if not sheet_found:
                ws = wb.worksheets[1]

            # 3. MAPEO DE COLUMNAS (Fila 4)
            # Vamos a recorrer hasta la columna 500 para encontrar todo
            amazon_cols = {}
            for c in range(1, 501):
                cell_val = ws.cell(row=4, column=c).value
                if cell_val:
                    amazon_cols[str(cell_val).strip()] = c

            if not amazon_cols:
                st.error("No se detectaron cabeceras en la Fila 4. Verifica que sea la pestaña correcta.")
            else:
                # 4. ESCRITURA (Fila 7)
                rows_written = 0
                for i, row_pim in df_pim.iterrows():
                    target_row = i + 7 
                    
                    # Variables
                    for amz_key, pim_col in final_mapping.items():
                        if amz_key in amazon_cols and pim_col in df_pim.columns:
                            # Escribimos directamente en la celda
                            ws.cell(row=target_row, column=amazon_cols[amz_key]).value = row_pim[pim_col]
                    
                    # Fijos
                    for amz_key, val in VALORES_FIJOS.items():
                        if amz_key in amazon_cols:
                            ws.cell(row=target_row, column=amazon_cols[amz_key]).value = val
                    
                    rows_written += 1

                # 5. GUARDAR Y DESCARGAR
                output = BytesIO()
                wb.save(output)
                st.success(f"✅ ¡Éxito! Se han rellenado {rows_written} filas.")
                st.download_button(
                    label="📥 Descargar Resultado Final",
                    data=output.getvalue(),
                    file_name="Amazon_Bulk_Fill_OK.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error técnico: {e}")