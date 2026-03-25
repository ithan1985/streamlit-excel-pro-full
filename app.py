import io
from typing import List

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Excel Explorer Pro",
    page_icon="📊",
    layout="wide"
)

st.title("Excel Explorer Pro")
st.caption("Carga cualquier Excel, explora hojas, busca, filtra y descarga resultados.")

# -----------------------------
# Funciones auxiliares
# -----------------------------
@st.cache_data(show_spinner=False)
def obtener_hojas(file_bytes: bytes) -> List[str]:
    """Devuelve la lista de hojas del archivo Excel."""
    excel_buffer = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(excel_buffer)
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def leer_hoja(file_bytes: bytes, hoja: str) -> pd.DataFrame:
    """Lee una hoja específica del Excel y la retorna como DataFrame."""
    excel_buffer = io.BytesIO(file_bytes)
    return pd.read_excel(excel_buffer, sheet_name=hoja)


def limpiar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Limpieza básica segura para exploración."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()

    return df


def aplicar_busqueda_global(df: pd.DataFrame, texto: str) -> pd.DataFrame:
    """Busca un texto en cualquier columna."""
    if not texto:
        return df

    mascara = df.astype(str).apply(
        lambda col: col.str.contains(texto, case=False, na=False)
    ).any(axis=1)

    return df[mascara]


def convertir_a_csv(df: pd.DataFrame) -> bytes:
    """Convierte un DataFrame a CSV descargable."""
    return df.to_csv(index=False).encode("utf-8")


def mostrar_metricas(df_original: pd.DataFrame, df_filtrado: pd.DataFrame, archivo: str, hoja: str) -> None:
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Archivo", archivo)
    c2.metric("Hoja", hoja)
    c3.metric("Filas originales", len(df_original))
    c4.metric("Filas filtradas", len(df_filtrado))
    c5.metric("Columnas", len(df_filtrado.columns))


def mostrar_resumen(df_filtrado: pd.DataFrame) -> None:
    st.subheader("Resumen del dataset")

    info_df = pd.DataFrame({
        "columna": df_filtrado.columns,
        "tipo": [str(t) for t in df_filtrado.dtypes.values],
        "nulos": [int(df_filtrado[c].isna().sum()) for c in df_filtrado.columns],
    })
    st.write("**Estructura y tipos de datos**")
    st.dataframe(info_df, use_container_width=True)

    columnas_numericas = df_filtrado.select_dtypes(include=["number"]).columns.tolist()
    columnas_texto = df_filtrado.select_dtypes(include=["object"]).columns.tolist()

    if columnas_numericas:
        st.write("**Estadísticas numéricas**")
        st.dataframe(df_filtrado[columnas_numericas].describe(), use_container_width=True)

        col_graf = st.selectbox(
            "Seleccione una columna numérica para gráfico",
            columnas_numericas,
            key="grafico_numerico"
        )
        st.bar_chart(df_filtrado[col_graf])

    if columnas_texto:
        st.write("**Frecuencia de columnas de texto**")
        col_freq = st.selectbox(
            "Seleccione una columna de texto para conteo",
            columnas_texto,
            key="frecuencia_texto"
        )
        frecuencia = df_filtrado[col_freq].value_counts(dropna=False).reset_index()
        frecuencia.columns = [col_freq, "cantidad"]
        st.dataframe(frecuencia, use_container_width=True)


# -----------------------------
# Interfaz principal
# -----------------------------
archivo = st.file_uploader(
    "Sube un archivo Excel",
    type=["xlsx", "xls", "xlsm", "xlsb", "ods"]
)

if archivo is None:
    st.info("Sube un archivo Excel para comenzar.")
    st.markdown("""
### Sugerencias de prueba
- Usa el archivo de ejemplo dentro de la carpeta `data/`.
- Prueba con varias hojas.
- Usa búsqueda global y filtros desde la barra lateral.
- Descarga el resultado filtrado en CSV.
""")
    st.stop()

try:
    contenido = archivo.getvalue()
    hojas = obtener_hojas(contenido)

    with st.sidebar:
        st.header("Configuración")
        hoja_seleccionada = st.selectbox("Hoja", hojas)

    df = leer_hoja(contenido, hoja_seleccionada)
    df = limpiar_dataframe(df)

    if df.empty:
        st.warning("La hoja seleccionada no contiene datos útiles después de la limpieza.")
        st.stop()

    df_filtrado = df.copy()

    with st.sidebar:
        st.header("Búsqueda y filtros")

        texto_busqueda = st.text_input("Buscar en todo el archivo")
        if texto_busqueda:
            df_filtrado = aplicar_busqueda_global(df_filtrado, texto_busqueda)

        columnas_texto = df_filtrado.select_dtypes(include=["object"]).columns.tolist()
        columnas_numericas = df_filtrado.select_dtypes(include=["number"]).columns.tolist()

        filtro_texto_col = st.selectbox(
            "Filtrar por columna de texto",
            ["Ninguna"] + columnas_texto
        )

        if filtro_texto_col != "Ninguna":
            opciones = sorted(
                df_filtrado[filtro_texto_col]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            valores_texto = st.multiselect(
                f"Valores en {filtro_texto_col}",
                options=opciones
            )

            if valores_texto:
                df_filtrado = df_filtrado[
                    df_filtrado[filtro_texto_col].astype(str).isin(valores_texto)
                ]

        filtro_num_col = st.selectbox(
            "Filtrar por columna numérica",
            ["Ninguna"] + columnas_numericas
        )

        if filtro_num_col != "Ninguna" and not df_filtrado.empty:
            min_val = float(df_filtrado[filtro_num_col].min())
            max_val = float(df_filtrado[filtro_num_col].max())

            rango = st.slider(
                f"Rango de {filtro_num_col}",
                min_value=min_val,
                max_value=max_val,
                value=(min_val, max_val)
            )

            df_filtrado = df_filtrado[
                (df_filtrado[filtro_num_col] >= rango[0]) &
                (df_filtrado[filtro_num_col] <= rango[1])
            ]

    mostrar_metricas(df, df_filtrado, archivo.name, hoja_seleccionada)

    tab1, tab2, tab3 = st.tabs(["Datos", "Resumen", "Descarga"])

    with tab1:
        st.subheader("Vista de datos")
        st.dataframe(df_filtrado, use_container_width=True)

    with tab2:
        mostrar_resumen(df_filtrado)

    with tab3:
        st.subheader("Descargar resultados")
        csv_data = convertir_a_csv(df_filtrado)
        st.download_button(
            label="Descargar CSV filtrado",
            data=csv_data,
            file_name="resultado_filtrado.csv",
            mime="text/csv"
        )

except Exception as e:
    st.error(f"Ocurrió un error al procesar el archivo: {e}")
