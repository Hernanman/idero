import hashlib
import tempfile
from pathlib import Path

import streamlit as st

from djim_core import DJIM_PARSER_VERSION, procesar_djim_web

st.set_page_config(
    page_title="DJIM Automatiza 360",
    page_icon="馃搫",
    layout="centered",
)

st.title("馃搫 DJIM Automatiza 360")
st.caption("Generador local de TXT DNRPA y Excel DJIM desde PDF ARCA-SIM. Sin APIs pagas.")

st.warning(
    "Esta versi贸n gratuita usa extracci贸n por texto y reglas. Funciona mejor con PDFs con texto seleccionable. "
    "Si el PDF es una imagen escaneada, puede no detectar todos los datos."
)

# Los archivos generados se guardan en session_state para que NO desaparezcan
# cuando se descarga TXT o Excel. Streamlit recarga la p谩gina al tocar botones,
# por eso no conviene depender de archivos temporales luego del procesamiento.
if "resultado_djim" not in st.session_state:
    st.session_state["resultado_djim"] = None

# Al actualizar el parser se descarta cualquier resultado generado con una
# versi贸n anterior, aunque el usuario mantenga cargado exactamente el mismo PDF.
resultado_guardado = st.session_state.get("resultado_djim")
if resultado_guardado and resultado_guardado.get("parser_version") != DJIM_PARSER_VERSION:
    st.session_state["resultado_djim"] = None

pdf_file = st.file_uploader("Sub铆 el PDF del despacho", type=["pdf"])
template_file = st.file_uploader("Template DJIM Excel opcional", type=["xlsx"])

# Vinculamos el resultado al PDF/template realmente cargados. As铆, al cambiar
# de despacho, desaparecen las descargas anteriores y nunca se entrega un TXT
# perteneciente a otro PDF que hubiera quedado guardado en session_state.
source_key = None
pdf_bytes = None
template_bytes = None
if pdf_file is not None:
    pdf_bytes = pdf_file.getvalue()
    template_bytes = template_file.getvalue() if template_file is not None else b""
    source_key = hashlib.sha256(
        pdf_bytes
        + b"|DJIM_TEMPLATE|"
        + template_bytes
        + b"|PARSER_VERSION|"
        + DJIM_PARSER_VERSION.encode("utf-8")
    ).hexdigest()

    resultado_anterior = st.session_state.get("resultado_djim")
    if resultado_anterior and resultado_anterior.get("source_key") != source_key:
        st.session_state["resultado_djim"] = None

procesar = st.button("Generar TXT / Excel", type="primary", disabled=pdf_file is None)

if procesar and pdf_file:
    with st.spinner("Procesando PDF y generando archivos..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)
                pdf_path = tmpdir_path / pdf_file.name
                pdf_path.write_bytes(pdf_bytes if pdf_bytes is not None else pdf_file.getvalue())

                template_path = None
                if template_file is not None:
                    template_path = tmpdir_path / template_file.name
                    template_path.write_bytes(template_bytes if template_bytes is not None else template_file.getvalue())

                result = procesar_djim_web(
                    pdf_path=str(pdf_path),
                    output_dir=str(tmpdir_path),
                    template_path=str(template_path) if template_path else None,
                )

                txt_path = Path(result["txt_path"])
                xlsx_path = Path(result["xlsx_path"]) if result.get("xlsx_path") else None

                # Guardamos bytes y nombres en memoria de sesi贸n.
                st.session_state["resultado_djim"] = {
                    "source_key": source_key,
                    "parser_version": DJIM_PARSER_VERSION,
                    "source_pdf_name": pdf_file.name,
                    "datos": result["datos"],
                    "campos_vacios": result.get("campos_vacios", []),
                    "txt_name": "DJIM_ELECTRONICA.txt",
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
        st.metric("Veh铆culos detectados", len(vehiculos))
    with col2:
        st.metric("Aduana", cab.get("aduana_nombre", ""))
        st.metric("Fecha oficializaci贸n", cab.get("fecha_oficializacion", ""))

    campos_vacios = resultado.get("campos_vacios", [])
    if campos_vacios:
        st.warning("Campos importantes no detectados autom谩ticamente. Revisalos antes de presentar:")
        st.write(campos_vacios)

    salida_invalida = not cab.get("nro_despacho_raw") or not cab.get("fecha_oficializacion")
    if salida_invalida:
        st.error(
            "No se habilita la descarga porque falta el n煤mero de despacho o la fecha de oficializaci贸n. "
            "Volv茅 a generar el archivo con esta versi贸n del parser."
        )

    with st.expander("Ver JSON extra铆do solo para control interno"):
        st.json(datos)

    st.subheader("Descargas")

    st.download_button(
        "猬囷笍 Descargar TXT DNRPA",
        data=resultado["txt_bytes"],
        file_name=resultado["txt_name"],
        mime="text/plain",
        key="download_txt",
        disabled=salida_invalida,
    )

    if resultado.get("xlsx_bytes"):
        st.download_button(
            "猬囷笍 Descargar Excel DJIM",
            data=resultado["xlsx_bytes"],
            file_name=resultado["xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_xlsx",
            disabled=salida_invalida,
        )
    else:
        st.info("No se gener贸 Excel porque no subiste template DJIM .xlsx.")

st.divider()
st.caption("Automatiza 360 路 Versi贸n sin IA/API 路 Revisi贸n manual recomendada antes de presentar.")
