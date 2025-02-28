import tempfile
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from zipfile import ZipFile
import os
import subprocess
import logging
from typing import Dict, Any, List, Optional

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="Generador de Certificados", page_icon="📄", layout="wide"
)


def validate_data(data: pd.DataFrame) -> bool:
    """Valida que el DataFrame contenga las columnas necesarias."""
    required_columns = ["Nombre Completo"]
    missing_columns = [col for col in required_columns if col not in data.columns]

    if missing_columns:
        st.error(
            f"Faltan las siguientes columnas requeridas: {', '.join(missing_columns)}"
        )
        return False
    return True


def convert_to_docx(input_path: str, output_dir: str) -> str:
    """Convierte un archivo .doc a .docx usando LibreOffice."""
    try:
        filename = os.path.basename(input_path)
        name_without_ext = os.path.splitext(filename)[0]
        output_path = os.path.join(output_dir, f"{name_without_ext}.docx")

        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "docx",
                input_path,
                "--outdir",
                output_dir,
            ],
            capture_output=True,
            check=False,
        )

        if result.returncode != 0:
            error_msg = result.stderr.decode()
            logger.error(f"Error al convertir .doc a .docx: {error_msg}")
            raise Exception(f"Error al convertir el archivo: {error_msg}")

        # Verificar que el archivo existe después de la conversión
        if not os.path.exists(output_path):
            raise Exception("El archivo convertido no se creó correctamente")

        return output_path

    except Exception as e:
        logger.exception("Error en la conversión de .doc a .docx")
        raise Exception(f"Error en la conversión del formato: {str(e)}")


def create_document(
    doc: DocxTemplate, row: pd.Series, output_dir: str
) -> Dict[str, str]:
    """Crea un documento individual y lo convierte a PDF."""
    try:
        nombre = row.get("Nombre Completo", "").title()
        context = {
            "nombre_completo": nombre,
            "cargo": row.get("Cargo", ""),
        }

        # Generar archivo Word
        docx_path = os.path.join(output_dir, f"Certificado - {nombre}.docx")
        doc.render(context)
        doc.save(docx_path)

        # Convertir a PDF
        pdf_path = os.path.join(output_dir, f"Certificado - {nombre}.pdf")
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                docx_path,
                "--outdir",
                output_dir,
            ],
            capture_output=True,
            check=False,
        )

        if result.returncode != 0:
            logger.error(f"Error al convertir a PDF: {result.stderr.decode()}")
            return {"status": "error", "message": f"Error al convertir {nombre} a PDF"}

        return {"status": "success", "file": pdf_path, "nombre": nombre}

    except Exception as e:
        logger.exception(
            f"Error al procesar documento para {row.get('Nombre Completo', '')}"
        )
        return {"status": "error", "message": str(e), "nombre": nombre}


def create_documents(
    template_path: str, data: pd.DataFrame, output_dir: str, progress_bar: Any
) -> List[Dict[str, str]]:
    """Procesa todos los documentos y muestra el progreso."""
    doc = DocxTemplate(template_path)
    total_files = len(data)
    results = []

    for i, row in data.iterrows():
        # Actualizar la barra de progreso
        progress_text = f"Generando certificado {i+1} de {total_files}"
        progress_bar.progress((i) / total_files, text=progress_text)

        # Crear documento
        result = create_document(doc, row, output_dir)
        results.append(result)

        if result["status"] == "error":
            st.warning(
                f"Problema al generar certificado para {result.get('nombre', 'un participante')}: {result.get('message', '')}"
            )

    progress_bar.progress(1.0, text="¡Proceso completado!")

    return results


def create_zip_file(results: List[Dict[str, str]], output_dir: str) -> Optional[str]:
    """Crea un archivo ZIP con los PDF generados exitosamente."""
    successful_files = [r["file"] for r in results if r["status"] == "success"]

    if not successful_files:
        st.error("No se pudieron generar certificados para crear un archivo ZIP.")
        return None

    zip_path = os.path.join(output_dir, "Certificados.zip")
    try:
        with ZipFile(zip_path, "w") as zipf:
            for file_path in successful_files:
                file_name = os.path.basename(file_path)
                zipf.write(file_path, arcname=file_name)

        return zip_path
    except Exception as e:
        logger.exception("Error al crear archivo ZIP")
        st.error(f"Error al crear archivo ZIP: {str(e)}")
        return None


def process_template_file(template_file, temp_dir: str) -> str:
    """Procesa el archivo de plantilla, convirtiendo de .doc a .docx si es necesario."""
    file_extension = os.path.splitext(template_file.name)[1].lower()
    original_path = os.path.join(temp_dir, f"template{file_extension}")

    with open(original_path, "wb") as f:
        f.write(template_file.getbuffer())

    # Si es .doc, convertir a .docx
    if file_extension == ".doc":
        st.info("Formato .doc detectado. Convirtiendo a .docx para compatibilidad...")
        try:
            docx_path = convert_to_docx(original_path, temp_dir)
            st.success("Conversión a .docx completada correctamente.")
            return docx_path
        except Exception as e:
            st.error(f"Error al convertir el archivo: {str(e)}")
            st.stop()
    else:
        # Si ya es .docx, usarlo directamente
        return original_path


def main():
    st.image("assets/background.jpg", use_container_width=True)
    st.markdown(
        "<h1 style='text-align: center;'>Generador de Certificados</h1>",
        unsafe_allow_html=True,
    )
    st.write(
        "Sube una plantilla base (.doc o .docx) y una lista de participantes (.xlsx, .csv) para generar certificados PDF."
    )

    # Contenedores de carga de archivos
    col1, col2 = st.columns(2)

    with col1:
        template_file = st.file_uploader(
            "Subir plantilla (.doc o .docx) 📄",
            type=["doc", "docx"],
            help='La plantilla debe contener el texto "{{ nombre_completo }}" donde iràn los datos de los participantes.',
        )
        if template_file:
            st.success(f"Plantilla '{template_file.name}' subida exitosamente.")

    with col2:
        data_file = st.file_uploader(
            "Subir datos (.xlsx, .csv) 📊",
            type=["xlsx", "csv"],
            help="El archivo debe contener la columna 'Nombre Completo'.",
        )
        if data_file:
            st.success(f"Datos '{data_file.name}' subidos exitosamente.")

    # Procesamiento de archivos
    if template_file and data_file:
        try:
            # Crear directorio temporal para los archivos
            with tempfile.TemporaryDirectory() as temp_dir:
                # Procesar el archivo de plantilla
                with st.spinner("Procesando plantilla..."):
                    template_path = process_template_file(template_file, temp_dir)

                # Cargar datos
                try:
                    with st.spinner("Cargando datos..."):
                        if data_file.name.endswith(".xlsx"):
                            data = pd.read_excel(data_file)
                        else:
                            data = pd.read_csv(data_file)
                except Exception as e:
                    st.error(f"Error al cargar el archivo de datos: {str(e)}")
                    st.stop()

                # Mostrar vista previa y validar datos
                st.subheader("Vista previa de los datos:")
                st.dataframe(data, height=200)

                if not validate_data(data):
                    st.stop()

                st.info(f"Se generarán certificados para {len(data)} participantes.")

                # Generar botón y manejo de la generación
                if st.button("Generar Certificados", type="primary"):
                    with st.spinner("Generando certificados..."):
                        # Crear barra de progreso
                        progress_bar = st.progress(0, text="Iniciando generación...")

                        # Generar documentos
                        results = create_documents(
                            template_path, data, temp_dir, progress_bar
                        )

                        # Estadísticas de resultados
                        successful = sum(1 for r in results if r["status"] == "success")
                        failed = len(results) - successful

                        # Crear ZIP con los archivos exitosos
                        if successful > 0:
                            zip_path = create_zip_file(results, temp_dir)

                            if zip_path:
                                st.success(
                                    f"✅ {successful} certificados generados exitosamente."
                                )
                                if failed > 0:
                                    st.warning(
                                        f"⚠️ {failed} certificados no pudieron ser generados."
                                    )

                                # Botón de descarga
                                with open(zip_path, "rb") as zipf:
                                    st.download_button(
                                        label="📥 Descargar todos los certificados",
                                        data=zipf,
                                        file_name="Certificados.zip",
                                        mime="application/zip",
                                    )
                        else:
                            st.error("No se pudo generar ningún certificado.")

        except Exception as e:
            logger.exception("Error en el procesamiento principal")
            st.error(f"Ocurrió un error inesperado: {str(e)}")


if __name__ == "__main__":
    main()
