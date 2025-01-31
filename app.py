import tempfile
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from zipfile import ZipFile
import os
import subprocess


def create_documents(template_path, data, output_dir):
    doc = DocxTemplate(template_path)

    for i, row in data.iterrows():
        context = {
            "nombre_completo": row.get("Nombre Completo", "").title(),
            "cargo": row.get("Cargo", ""),
        }
        output_path = os.path.join(
            output_dir, f"Certificado - {row.get('Nombre Completo', '').title()}.docx"
        )
        doc.render(context)
        doc.save(output_path)
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                output_path,
                "--outdir",
                output_dir,
            ],
            check=True,
        )


def main():
    if "toast" not in st.session_state:
        st.session_state.toast = False

    if not st.session_state.toast:
        st.toast("Debe agregar {{ nombre_completo }} en la plantilla .docx", icon="❗")
        st.session_state.toast = True

    st.image("assets/background.jpg", use_container_width=True)
    st.title("Generador de Certificados")
    st.write(
        "Sube una plantilla base y una lista de participantes en formato Excel o CSV."
    )

    # template_file = st.file_uploader("Sube la plantilla base (.docx)", type=["docx"])
    # data_file = st.file_uploader(
    #     "Sube la lista de participantes (.xlsx, .csv)", type=["xlsx", "csv"]
    # )
    col1, col2 = st.columns(2)

    with col1:
        template_file = st.file_uploader("Subir plantilla (.docx) 📄 📃", type=["docx"])
        if template_file:
            st.success("Plantilla subida exitosamente.")

    with col2:
        data_file = st.file_uploader(
            "Subir datos (.xlsx, .csv) 💹 📈", type=["xlsx", "csv"]
        )
        if data_file:
            st.success("Datos subidos exitosamente.")

    if template_file and data_file:
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = os.path.join(temp_dir, "template.docx")
            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())

            if data_file.name.endswith(".xlsx"):
                data = pd.read_excel(data_file)
            else:
                data = pd.read_csv(data_file)

            st.write("Vista previa de los datos:")
            st.dataframe(data)
            st.warning(
                "Asegúrese que exista la columna 'Nombre Completo' para generar los certificados."
            )

            if st.button("Generar Certificados"):
                create_documents(template_path, data, temp_dir)

                zip_path = os.path.join(temp_dir, "Certificados.zip")
                with ZipFile(zip_path, "w") as zipf:
                    for file in os.listdir(temp_dir):
                        if file.endswith(".pdf"):
                            zipf.write(os.path.join(temp_dir, file), arcname=file)

                st.success("Documentos generados exitosamente.")
                with open(zip_path, "rb") as zipf:
                    st.download_button(
                        label="Descargar todos los documentos",
                        data=zipf,
                        file_name="Certificados.zip",
                        mime="application/zip",
                    )


if __name__ == "__main__":
    main()
