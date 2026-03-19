import streamlit as st
import pandas as pd
from bbva_extractor import extraer_movimientos_desde_pdf, generar_excel_movimientos
import io
import tempfile
import os

st.set_page_config(page_title="Extractor BBVA", page_icon="💳", layout="centered")

st.title("💳 Extractor BBVA")
st.markdown("**Extrae movimientos de tu estado de cuenta PDF a Excel fácilmente**")

# Zona de subida de PDF
uploaded_file = st.file_uploader("Sube tu archivo PDF del estado de cuenta BBVA", type=["pdf"], help="Arrastra y suelta o haz clic para seleccionar")

if uploaded_file is not None:
    st.success("PDF cargado exitosamente")

    # Botón para extraer
    if st.button("🚀 EXTRAER DATOS", type="primary"):
        with st.spinner("Procesando PDF... Esto puede tomar unos minutos"):
            progress_bar = st.progress(0)
            progress_bar.progress(10)

            # Guardar temporalmente el PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                pdf_path = tmp_file.name

            progress_bar.progress(30)

            try:
                # Extraer movimientos
                movimientos, info = extraer_movimientos_desde_pdf(pdf_path, log=st.write)
                progress_bar.progress(70)

                if movimientos:
                    st.success(f"✅ Se encontraron {len(movimientos)} movimientos")

                    # Mostrar preview de la tabla
                    df = pd.DataFrame(movimientos)
                    st.subheader("📊 Preview de Movimientos")
                    st.dataframe(df.head(20), use_container_width=True)

                    # Generar Excel en memoria
                    excel_buffer = io.BytesIO()
                    generar_excel_movimientos(movimientos, info, excel_buffer, log=st.write)
                    progress_bar.progress(100)

                    # Botón de descarga
                    st.download_button(
                        label="📥 DESCARGAR EXCEL",
                        data=excel_buffer.getvalue(),
                        file_name="movimientos_bbva.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.warning("⚠ No se encontraron movimientos en el PDF")

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
            finally:
                # Limpiar archivo temporal
                os.unlink(pdf_path)
                progress_bar.empty()

st.markdown("---")
st.markdown("**Instrucciones:** Sube tu PDF, haz clic en EXTRAER DATOS, espera el procesamiento y descarga el Excel.")
st.markdown("Funciona en cualquier dispositivo sin instalación.")