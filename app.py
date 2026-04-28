import tempfile
from pathlib import Path

import streamlit as st

from djim_core import procesar_djim_web

st.set_page_config(
    page_title="Automatiza 360",
    page_icon="📄",
    layout="centered",
)

st.title("📄 DJIM Automatiza 360")
st.caption("Generador local de TXT DNRPA y Excel DJIM desde PDF ARCA-SIM.")

st.warning(
    "Esta versión gratuita usa extracción por texto y reglas. Funciona mejor con PDFs con texto seleccionable. "
    "Si el PDF es una imagen escaneada, puede no detectar todos los datos."
)

# Los archivos generados se guardan en session_state para que NO desaparezcan
# cuando se descarga TXT o Excel. Streamlit recarga la página al tocar botones,
# por eso no conviene depender de archivos temporales luego del procesamiento.
if "resultado_djim" not in st.session_state:
    st.session_state["resultado_djim"] = None

pdf_file = st.file_uploader("Subí el PDF del despacho", type=["pdf"])
template_file = st.file_uploader("Template DJIM Excel opcional", type=["xlsx"])

procesar = st.button("Generar TXT / Excel", type="primary", disabled=pdf_file is None)

if procesar and pdf_file:
    with st.spinner("Procesando PDF y generando archivos..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)
                pdf_path = tmpdir_path / pdf_file.name
                pdf_path.write_bytes(pdf_file.getbuffer())

                template_path = None
                if template_file is not None:
                    template_path = tmpdir_path / template_file.name
                    template_path.write_bytes(template_file.getbuffer())

                result = procesar_djim_web(
                    pdf_path=str(pdf_path),
                    output_dir=str(tmpdir_path),
                    template_path=str(template_path) if template_path else None,
                )

                txt_path = Path(result["txt_path"])
                xlsx_path = Path(result["xlsx_path"]) if result.get("xlsx_path") else None

                # Guardamos bytes y nombres en memoria de sesión.
                st.session_state["resultado_djim"] = {
                    "datos": result["datos"],
                    "campos_vacios": result.get("campos_vacios", []),
                    "txt_name": txt_path.name,
                    "txt_bytes": txt_path.read_bytes(),
                    "xlsx_name": xlsx_path.name if xlsx_path else None,
                    "xlsx_bytes": xlsx_path.read_bytes() if xlsx_path else None,
                }

            st.success("Proceso completado. Los archivos quedan disponibles para descargar abajo.")

        except Exception as e:
            st.session_state["resultado_djim"] = None
            st.error("No se pudo procesar el PDF.")
            st.exception(e)

resultado = st.session_state.get("resultado_djim")

if resultado:
    datos = resultado["datos"]
    cab = datos.get("cabecera", {})
    vehiculos = datos.get("vehiculos", [])

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Despacho", cab.get("nro_despacho_raw", ""))
        st.metric("Vehículos detectados", len(vehiculos))
    with col2:
        st.metric("Aduana", cab.get("aduana_nombre", ""))
        st.metric("Fecha oficialización", cab.get("fecha_oficializacion", ""))

    campos_vacios = resultado.get("campos_vacios", [])
    if campos_vacios:
        st.warning("Campos importantes no detectados automáticamente. Revisalos antes de presentar:")
        st.write(campos_vacios)

    with st.expander("Ver JSON extraído solo para control interno"):
        st.json(datos)

    st.subheader("Descargas")

    st.download_button(
        "⬇️ Descargar TXT DNRPA",
        data=resultado["txt_bytes"],
        file_name=resultado["txt_name"],
        mime="text/plain",
        key="download_txt",
    )

    if resultado.get("xlsx_bytes"):
        st.download_button(
            "⬇️ Descargar Excel DJIM",
            data=resultado["xlsx_bytes"],
            file_name=resultado["xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_xlsx",
        )
    else:
        st.info("No se generó Excel porque no subiste template DJIM .xlsx.")

st.divider()
st.caption("Automatiza 360 · Versión sin IA/API · Revisión manual recomendada antes de presentar.")
