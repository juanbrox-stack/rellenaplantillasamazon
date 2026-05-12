import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Amazon Filler Pro", layout="wide")

st.title("📦 Amazon Template Automator")

# --- CARGA DE ARCHIVOS ---
with st.sidebar:
    st.header("1. Configuración de Mapeo")
    config_file = st.file_uploader("Subir Amazon-PIM.xlsx", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    st.header("2. Fichero PIM")
    pim_file = st.file_uploader("Subir FRIGOS.xlsx", type=["xlsx", "csv"])

with col2:
    st.header("3. Plantilla Amazon")
    amazon_template = st.file_uploader("Subir Plantilla Amazon (Refrigerators...)", type=["xlsx"])

if st.button("🚀 Iniciar Carga Masiva"):
    if not (config_file and pim_file and amazon_template):
        st.error("Faltan archivos por subir.")
    else:
        try:
            # 1. LEER CONFIGURACIÓN (Mapeo y Fijos)
            # Aseguramos nombres de columnas limpios
            df_map = pd.read_excel(config_file, sheet_name=0)
            df_fixed = pd.read_excel(config_file, sheet_name=1)

            # 2. LEER PIM (FRIGOS)
            df_pim = pd.read_excel(pim_file) if pim_file.name.endswith('.xlsx') else pd.read_csv(pim_file)

            # 3. LEER PLANTILLA AMAZON (Pestaña 2: 'Template')
            # Amazon usa la pestaña 2 para los datos. Habitualmente se llama 'Template'
            xl = pd.ExcelFile(amazon_template)
            nombre_pestana_destino = xl.sheet_names[1] # Forzamos la segunda pestaña
            
            # Guardamos las 3 filas de encabezado técnico
            headers = pd.read_excel(amazon_template, sheet_name=nombre_pestana_destino, header=None, nrows=3)
            # Obtenemos los nombres de las columnas técnicas (fila 3, índice 2)
            amazon_cols = headers.iloc[2].tolist()

            # 4. CREAR DATAFRAME DE TRABAJO
            num_rows = len(df_pim)
            # Creamos un DF vacío con las columnas exactas de Amazon
            df_final = pd.DataFrame(columns=amazon_cols)

            # 5. PROCESAR VALORES FIJOS (Hoja 2)
            # Usamos iloc para evitar errores si los nombres de columna en el mapeo fallan
            for _, row in df_fixed.iterrows():
                campo_amazon = str(row.iloc[0]).strip()
                valor_fijo = row.iloc[1]
                if campo_amazon in df_final.columns:
                    df_final[campo_amazon] = [valor_fijo] * num_rows

            # 6. PROCESAR MAPEO (Hoja 1)
            for _, row in df_map.iterrows():
                col_amazon = str(row['AMAZON']).strip()
                col_pim = str(row['PIM']).strip()
                
                if col_amazon in df_final.columns and col_pim in df_pim.columns:
                    df_final[col_amazon] = df_pim[col_pim].values

            # 7. RECONSTRUCCIÓN FINAL
            # Convertimos los encabezados a DF para unirlos
            df_headers = pd.DataFrame(headers.values, columns=amazon_cols)
            # Concatenamos: Encabezados + Datos nuevos
            df_completo = pd.concat([df_headers, df_final], ignore_index=True)

            # 8. GENERAR EXCEL
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Escribimos en la pestaña 'Template'
                df_completo.to_excel(writer, index=False, header=False, sheet_name=nombre_pestana_destino)
            
            st.success(f"✅ ¡Hecho! {num_rows} productos procesados.")
            st.download_button(
                label="⬇️ Descargar Fichero Final",
                data=output.getvalue(),
                file_name="Amazon_Carga_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error detallado: {e}")
            st.info("Asegúrate de que en el archivo de mapeo las columnas se llamen 'AMAZON' y 'PIM'.")