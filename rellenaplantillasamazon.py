import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Amazon Bulk Upload", layout="wide")

st.title("📦 Amazon Inventory Automator")
st.markdown("Automatiza el rellenado de plantillas de Amazon usando datos de PIM y reglas de mapeo.")

# --- CARGA DE ARCHIVOS ---
with st.sidebar:
    st.header("Configuración de Mapeo")
    # Aquí el usuario sube el archivo con Hoja1 (Mapeo) y Hoja2 (Fijos)
    config_file = st.file_uploader("Subir Amazon-PIM.xlsx", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    pim_file = st.file_uploader("1. Subir Datos Origen (PIM)", type=["xlsx", "csv"])
with col2:
    amazon_template = st.file_uploader("2. Subir Plantilla Amazon (Vacía)", type=["xlsx"])

if st.button("🚀 Generar Fichero de Carga Masiva"):
    if not config_file or not pim_file or not amazon_template:
        st.error("Por favor, sube los tres archivos para continuar.")
    else:
        try:
            # 1. LEER CONFIGURACIÓN
            # Hoja 1: Columnas AMAZON vs PIM | Hoja 2: Valores AUTOCOMPLETAR
            df_map = pd.read_excel(config_file, sheet_name=0) 
            df_fixed = pd.read_excel(config_file, sheet_name=1)

            # 2. LEER PIM
            if pim_file.name.endswith('.csv'):
                df_pim = pd.read_csv(pim_file)
            else:
                df_pim = pd.read_excel(pim_file)

            # 3. LEER PLANTILLA AMAZON
            # Importante: Amazon tiene encabezados en las filas 1, 2 y 3.
            # Leemos la plantilla entera para conservar los encabezados exactos.
            df_template_headers = pd.read_excel(amazon_template, header=None, nrows=3)
            df_template_cols = pd.read_excel(amazon_template, header=2) # La fila 3 (índice 2) tiene los nombres técnicos
            
            # Crear DataFrame final basado en la estructura de Amazon
            df_final = pd.DataFrame(columns=df_template_cols.columns)

            # --- PROCESAMIENTO ---
            num_items = len(df_pim)
            st.info(f"Procesando {num_items} productos...")

            # A. Rellenar Valores Fijos (Hoja 2)
            # Suponiendo columnas: "CAMPO" y "VALOR"
            for _, row in df_fixed.iterrows():
                col_name = str(row[0]).strip()
                val_fixed = row[1]
                if col_name in df_final.columns:
                    df_final[col_name] = [val_fixed] * num_items

            # B. Mapear Valores desde PIM (Hoja 1)
            # Suponiendo columnas: "AMAZON" y "PIM"
            for _, row in df_map.iterrows():
                col_amazon = str(row['AMAZON']).strip()
                col_pim = str(row['PIM']).strip()
                
                if col_amazon in df_final.columns and col_pim in df_pim.columns:
                    df_final[col_amazon] = df_pim[col_pim].values

            # --- RECONSTRUCCIÓN DEL ARCHIVO FINAL ---
            # Unimos los 3 encabezados originales con los nuevos datos
            # Convertimos los encabezados a DataFrame para concatenar
            final_output = pd.concat([df_template_headers, df_final], ignore_index=True)

            # --- DESCARGA ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Escribimos sin encabezados de pandas porque ya los incluimos en el concat
                final_output.to_excel(writer, index=False, header=False, sheet_name='Template')
            
            st.success("✅ Archivo generado con éxito.")
            st.download_button(
                label="⬇️ Descargar Plantilla Amazon Cumplimentada",
                data=output.getvalue(),
                file_name="Carga_Masiva_Amazon_LISTO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Se produjo un error: {e}")