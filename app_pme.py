import streamlit as st
import pandas as pd
import os
from PyPDF2 import PdfReader, PdfWriter
from rapidfuzz import fuzz

# === Rutas ===
ruta_script = os.path.abspath(__file__)
carpeta_script = os.path.dirname(ruta_script)
ruta_csv = os.path.join(carpeta_script, "acciones_indexadas.csv")
ruta_pdf = r"C:\Users\Felipe.Monjes\OneDrive - Corp.Mun. de Ed.Cultura y Recreación de La Florida\PME - SEP\2025\PME"

# === Cargar CSV ===
try:
    df = pd.read_csv(ruta_csv)
except FileNotFoundError:
    st.error("❌ No se encontró el archivo acciones_indexadas.csv.")
    st.stop()

st.set_page_config("Buscador de Acciones PME", layout="wide")
st.title("🔍 Buscador Inteligente de Acciones PME")

# === Tabla resumen de acciones por colegio ===
st.subheader("📊 Resumen general de acciones encontradas")
resumen = df.groupby("Establecimiento")["Acción"].nunique().reset_index()
resumen.columns = ["Establecimiento", "Cantidad de Acciones"]
st.dataframe(resumen.sort_values("Establecimiento"))

# === Búsqueda inteligente de acciones ===
st.subheader("🔎 Buscar acciones por palabra clave")
texto_busqueda = st.text_input("Escribe una palabra clave (ej. retroalimentación, asistencia)")

if texto_busqueda:
    acciones_unicas = df["Acción"].unique()
    acciones_similares = [a for a in acciones_unicas if fuzz.partial_ratio(texto_busqueda.lower(), a.lower()) > 85]

    if acciones_similares:
        st.markdown(f"✅ Se encontraron **{len(acciones_similares)}** acción(es) similares.")
        acciones_seleccionadas = st.multiselect(
            "Selecciona las acciones a analizar:",
            options=acciones_similares,
            default=acciones_similares
        )
    else:
        st.warning("⚠️ No se encontraron acciones similares.")
        acciones_seleccionadas = []
else:
    acciones_seleccionadas = []

# === Filtro por acción seleccionada ===
if acciones_seleccionadas:
    filtro = df[df["Acción"].isin(acciones_seleccionadas)]
    colegios_disponibles = sorted(filtro["Establecimiento"].unique())

    st.markdown(f"📌 Las acciones seleccionadas están presentes en **{len(colegios_disponibles)}** colegio(s).")

    seleccionar_todos = st.checkbox("Seleccionar todos los colegios", value=True)
    colegios_seleccionados = st.multiselect(
        "Selecciona los colegios:",
        options=colegios_disponibles,
        default=colegios_disponibles if seleccionar_todos else []
    )

    # === Tabla filtrada (vista previa) ===
    st.subheader("👁 Vista previa de coincidencias")
    st.dataframe(filtro[filtro["Establecimiento"].isin(colegios_seleccionados)].sort_values(["Establecimiento", "Acción"]))

    # === Exportar a Excel ===
    st.subheader("📥 Exportar coincidencias a Excel")
    df_excel = filtro[filtro["Establecimiento"].isin(colegios_seleccionados)]
    excel_path = os.path.join(carpeta_script, "acciones_filtradas.xlsx")
    df_excel.to_excel(excel_path, index=False)
    with open(excel_path, "rb") as f:
        st.download_button("📊 Descargar Excel", f, file_name="acciones_filtradas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # === Exportar a PDF ===
    st.subheader("📄 Exportar páginas PDF con las acciones")
    if st.button("Generar PDF con páginas seleccionadas"):
        writer = PdfWriter()
        errores = []

        for colegio in colegios_seleccionados:
            filas = filtro[filtro["Establecimiento"] == colegio]
            for _, fila in filas.iterrows():
                archivo_pdf = os.path.join(ruta_pdf, fila["Documento PDF"])
                pagina = int(fila["Página"]) - 1
                try:
                    reader = PdfReader(archivo_pdf)
                    writer.add_page(reader.pages[pagina])
                except Exception as e:
                    errores.append(f"{fila['Documento PDF']} (pág {pagina+1}): {e}")

        if writer.pages:
            output_path = os.path.join(carpeta_script, "acciones_seleccionadas.pdf")
            with open(output_path, "wb") as f:
                writer.write(f)
            with open(output_path, "rb") as f:
                st.success("✅ PDF generado con éxito.")
                st.download_button("📥 Descargar PDF", f, file_name="acciones_seleccionadas.pdf", mime="application/pdf")

        if errores:
            st.warning("⚠️ Algunos errores al procesar:")
            for err in errores:
                st.text(f"• {err}")
