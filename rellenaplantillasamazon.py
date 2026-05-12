import streamlit as st
import pandas as pd
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Amazon Template Filler", layout="wide")

st.title("📦 Automatización de Plantillas Amazon")
st.markdown("""
Esta herramienta cruza los datos de tu **PIM** con la **Plantilla de Amazon** usando las reglas de mapeo y valores fijos proporcionadas.
""")

# --- SECCIÓN DE CARGA DE ARCHIVOS ---
st.sidebar.header("1. Configuración")
config_file = st.sidebar.file_uploader("Subir Amazon-PIM.xlsx (Mapeo y Fijos)", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    st.header("2. Datos Origen")
    pim_file = st.file_uploader("Subir Fichero PIM", type=["xlsx", "csv"])

with col2:
    st.header("3. Destino")
    amazon_template = st.file_uploader("Subir Plantilla Amazon Vacía", type=["xlsx"])

# --- LÓGICA DE PROCESAMIENTO ---
if st.button("Generar Plantilla Cumplimentada"):
    if not (config_file and pim_file and amazon_template):
        st.warning("Por favor, asegúrate de subir todos los archivos requeridos.")
    else:
        try:
            # A. Leer reglas de mapeo (Hojas del archivo de configuración)
            # Usamos los nombres de hoja que indicaste
            df_map = pd.read_excel(config_file, sheet_name=0)   # Hoja 1: AMAZON, PIM
            df_fixed = pd.read_excel(config_file, sheet_name=1) # Hoja 2: AUTOCOMPLETAR, Valor

            # B. Leer datos del PIM
            if pim_file.name.endswith('.csv'):
                df_pim = pd.read_csv(pim_file)
            else:
                df_pim = pd.read_excel(pim_file)

            # C. Leer Plantilla de Amazon conservando las 3 filas de cabecera
            # Leemos las 3 primeras filas para el encabezado final
            headers = pd.read_excel(amazon_template, header=None, nrows=3)
            # Leemos la estructura de columnas (fila 3 es el nombre técnico)
            df_structure = pd.read_excel(amazon_template, header=2)
            
            # Crear DataFrame vacío con las columnas exactas de Amazon
            df_final = pd.DataFrame(columns=df_structure.columns)

            # --- PROCESAMIENTO DE DATOS ---
            num_rows = len(df_pim)
            
            # 1. Aplicar Valores Fijos (Hoja 2)
            # La Hoja 2 tiene la columna 'AUTOCOMPLETAR' con el nombre del campo técnico
            for _, row in df_fixed.iterrows():
                campo_tecnico = str(row[0]).strip()
                valor_fijo = row[1]
                if campo_tecnico in df_final.columns:
                    df_final[campo_tecnico] = [valor_fijo] * num_rows

            # 2. Aplicar Mapeo Dinámico (Hoja 1)
            # La Hoja 1 tiene 'AMAZON' (destino) y 'PIM' (origen)
            for _, row in df_map.iterrows():
                col_amazon = str(row['AMAZON']).strip()
                col_pim = str(row['PIM']).strip()
                
                if col_amazon in df_final.columns and col_pim in df_pim.columns:
                    df_final[col_amazon] = df_pim[col_pim].values

            # --- RECONSTRUCCIÓN DEL EXCEL ---
            # Combinamos los encabezados originales con los nuevos datos
            # Convertimos headers a DataFrame con las mismas columnas para que encajen
            headers.columns = df_final.columns
            result_df = pd.concat([headers, df_final], ignore_index=True)

            # Exportar a memoria
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Escribimos sin encabezado de Pandas porque ya están en el DataFrame
                result_df.to_excel(writer, index=False, header=False, sheet_name='Template')
            
            st.success(f"✅ ¡Éxito! Se han procesado {num_rows} productos.")
            st.download_button(
                label="📥 Descargar Plantilla Lista",
                data=output.getvalue(),
                file_name="Amazon_Bulk_Load_READY.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Se ha producido un error durante el procesamiento: {e}")