    import streamlit as st
import os
import tempfile
import zipfile
import shutil
import sys
from pathlib import Path
from datetime import datetime
import subprocess

# Dependencias Pure Python para conversión
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from docx import Document 

# Configuración de la página
st.set_page_config(
    page_title="Conversor PDF Web - Streamlit Cloud",
    page_icon="☁️",
    layout="wide"
)

def main():
    st.title("☁️ CONVERSOR PDF WEB - STREAMLIT CLOUD")
    st.markdown("**(Sin dependencia de LibreOffice)**")
    st.markdown("**@Cristobal Almagro**")
    st.markdown("---")
    
    # Sidebar con configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        st.subheader("Formatos a convertir (Soporte Cloud):")
        # SOLO DOCX y TXT son viables sin LibreOffice en Cloud
        doc = st.checkbox("📄 Word (.docx)", value=True)
        txt = st.checkbox("📝 Texto (.txt)", value=True)
        
        # Otros formatos deshabilitados para Streamlit Cloud
        st.markdown("*(.doc, .rtf, .odt están deshabilitados. Su conversión requiere LibreOffice.)*")
        
        st.subheader("📁 Opciones de entrada:")
        subir_zip = st.checkbox("📦 Permitir subir carpetas (ZIP)", value=True)
        buscar_subcarpetas = st.checkbox("🔍 Buscar en subcarpetas", value=True)
        
        # Botón para verificar LibreOffice (Ahora solo informa)
        if st.button("🔍 Verificar Entorno"):
            st.info("🌐 **Modo Streamlit Cloud Detectado**")
            st.success("✅ La aplicación está configurada para usar librerías Pure Python (reportlab, python-docx).")
        
        st.markdown("---")
        
        # BOTÓN SALIR en el sidebar
        st.markdown("---")
        if st.button("🔒 SALIR", type="secondary", use_container_width=True):
            st.success("👋 ¡Hasta pronto! Cerrando la aplicación...")
            import time
            time.sleep(2)
            sys.exit()
        
        st.info("💡 Puedes subir archivos individuales o carpetas completas (ZIP)")
    
    # Área principal
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 Subir Archivos o Carpetas")
        
        # Determinar tipos de archivo permitidos
        allowed_types = []
        if doc: allowed_types.append('docx')
        if txt: allowed_types.append('txt')
        
        if not allowed_types:
            st.warning("⚠️ Selecciona al menos un tipo de archivo en la configuración (.docx o .txt)")
            return
        
        # Subida de archivos individuales
        uploaded_files = st.file_uploader(
            "📄 Archivos individuales",
            type=allowed_types,
            accept_multiple_files=True,
            help=f"Formatos permitidos en Streamlit Cloud: {', '.join(allowed_types)}"
        )
        
        # Subida de carpetas ZIP
        if subir_zip:
            st.markdown("---")
            uploaded_zip = st.file_uploader(
                "📦 Carpeta completa (archivo ZIP)",
                type=['zip'],
                help="Sube un archivo ZIP que contenga los documentos a convertir"
            )
        else:
            uploaded_zip = None
    
    with col2:
        st.header("📊 Control")
        
        total_files = len(uploaded_files) if uploaded_files else 0
        if uploaded_zip:
            total_files += 1 
        
        if total_files > 0:
            st.success(f"📦 {total_files} elementos listos para procesar")
            
            if st.button("🚀 INICIAR CONVERSIÓN", type="primary", use_container_width=True):
                process_all_files(uploaded_files, uploaded_zip, buscar_subcarpetas)
        else:
            st.info("⏳ Esperando archivos...")
        
        # Botón SALIR también en el área principal
        st.markdown("---")
        if st.button("🔒 CERRAR APLICACIÓN", type="secondary", use_container_width=True):
            st.success("👋 ¡Gracias por usar el Conversor PDF! Cerrando...")
            import time
            time.sleep(2)
            sys.exit()
    
    # Información adicional
    with st.expander("ℹ️ Información importante"):
        st.write("""
        Esta versión está optimizada para **Streamlit Cloud** al eliminar la dependencia de LibreOffice.
        
        **Formatos soportados:**
        - ✅ **.docx** (Word moderno) - via `python-docx` y `reportlab`
        - ✅ **.txt** (Texto) - via `reportlab`
        
        **Formatos no soportados en Cloud (requieren LibreOffice):**
        - ❌ .doc (Word antiguo)
        - ❌ .rtf (Texto enriquecido)
        - ❌ .odt (OpenDocument)
        """)

# --- Funciones de Conversión ---

def convert_txt(input_path, output_path):
    """Convierte archivo TXT a PDF usando reportlab"""
    try:
        # Leer archivo con diferentes codificaciones
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        content = None
        
        for encoding in encodings:
            try:
                with open(input_path, 'r', encoding=encoding) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            st.error("No se pudo leer el archivo TXT con ninguna codificación común")
            return False
        
        # Crear PDF
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        
        # Formatear texto para ReportLab
        formatted_text = content.replace('\n', '<br/>').replace('\t', '    ')
        story = [Paragraph(formatted_text, styles['Normal'])]
        
        doc.build(story)
        
        return os.path.exists(output_path) and os.path.getsize(output_path) > 0
        
    except Exception as e:
        # Usar st.exception para mostrar detalles
        st.exception(e)
        st.error(f"Error en conversión TXT: {e}")
        return False

def convert_docx(input_path, output_path):
    """Convierte archivo DOCX a PDF usando python-docx y reportlab"""
    try:
        # 1. Leer el documento DOCX
        document = Document(input_path)
        styles = getSampleStyleSheet()
        story = []
        
        # 2. Iterar sobre párrafos y estilos para construir el 'story' de ReportLab
        for paragraph in document.paragraphs:
            text = paragraph.text
            style_name = paragraph.style.name.lower()

            # Mapeo simple de estilos de Word a estilos de ReportLab
            if 'heading 1' in style_name:
                style = styles['Heading1']
            elif 'heading 2' in style_name:
                style = styles['Heading2']
            elif 'heading 3' in style_name:
                style = styles['Heading3']
            else:
                style = styles['Normal']
            
            # Crear el elemento Párrafo
            if text.strip():
                # ReportLab necesita reemplazar saltos de línea con <br/> en HTML-like markup
                formatted_text = text.replace('\n', '<br/>')
                story.append(Paragraph(formatted_text, style))
            
            # Agregar un espacio después del párrafo para mejorar la legibilidad
            story.append(Spacer(1, 6)) 
        
        # 3. Construir el PDF
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        doc.build(story)
        
        return os.path.exists(output_path) and os.path.getsize(output_path) > 0
        
    except Exception as e:
        st.exception(e)
        st.error(f"Error en conversión DOCX: {e}")
        return False

def convert_to_pdf(file_info, output_filename):
    """Convierte un archivo a PDF usando el nombre especificado y librerías Pure Python"""
    try:
        # Guardar archivo temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file_info['name']).suffix) as tmp_input:
            tmp_input.write(file_info['content'])
            input_path = tmp_input.name
        
        # Archivo de salida con nombre específico
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        
        extension = file_info['extension']
        success = False
        
        if extension == '.txt':
            success = convert_txt(input_path, output_path)
        elif extension == '.docx':
            success = convert_docx(input_path, output_path)
        else:
            # Informar que este formato no está soportado sin LibreOffice
            st.warning(f"⚠️ Formato {extension} no soportado en Streamlit Cloud.")
            success = False
        
        # Limpiar archivo temporal de entrada
        if os.path.exists(input_path):
            os.unlink(input_path)
        
        return output_path if success else None
        
    except Exception as e:
        st.error(f"Error en conversión de {file_info['name']}: {e}")
        return None

# --- Funciones de Flujo (sin cambios mayores) ---

def process_all_files(uploaded_files, uploaded_zip, buscar_subcarpetas):
    """Procesa tanto archivos individuales como ZIPs"""
    
    all_files_to_process = []
    
    # Procesar archivos individuales
    if uploaded_files:
        for uploaded_file in uploaded_files:
            all_files_to_process.append({
                'name': uploaded_file.name,
                'content': uploaded_file.getvalue(),
                'extension': Path(uploaded_file.name).suffix.lower()
            })
    
    # Procesar archivo ZIP
    if uploaded_zip:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_zip:
            tmp_zip.write(uploaded_zip.getvalue())
            zip_path = tmp_zip.name
        
        try:
            zip_files = extract_and_filter_zip(zip_path, buscar_subcarpetas)
            all_files_to_process.extend(zip_files)
        finally:
            if os.path.exists(zip_path):
                os.unlink(zip_path)
    
    if not all_files_to_process:
        st.error("❌ No se encontraron archivos para procesar")
        return
    
    # Procesar todos los archivos
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    converted_files = [] 
    log_messages = []
    
    # Área de log
    log_container = st.container()
    with log_container:
        st.subheader("📝 Registro de Actividad")
        log_placeholder = st.empty()
    
    for i, file_info in enumerate(all_files_to_process):
        # Filtrar solo extensiones soportadas
        if file_info['extension'] not in ['.txt', '.docx']:
             log_messages.append(f"[{datetime.now().strftime('%H:%M:%S')}] 🛑 Ignorando: {file_info['name']} (Formato no soportado en Cloud)")
             log_placeholder.text_area("", "\n".join(log_messages), height=200, key=f"log_skip_{i}")
             continue
             
        # Actualizar progreso
        progress = (i + 1) / len(all_files_to_process)
        progress_bar.progress(progress)
        status_text.text(f"Procesando: {file_info['name']} ({i+1}/{len(all_files_to_process)})")
        
        # Log
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_messages.append(f"[{timestamp}] 🔄 Convirtiendo: {file_info['name']}")
        log_placeholder.text_area("", "\n".join(log_messages), height=200, key=f"log_{i}")
        
        try:
            # Conversión
            with st.spinner(f"Convirtiendo {file_info['name']}..."):
                original_name = Path(file_info['name']).stem 
                pdf_filename = f"{original_name}.pdf"
                
                # Se llama a la función modificada que no usa LibreOffice
                pdf_path = convert_to_pdf(file_info, pdf_filename)
                
                if pdf_path and os.path.exists(pdf_path):
                    converted_files.append((pdf_filename, pdf_path))
                    log_messages.append(f"[{timestamp}] ✅ Convertido: {file_info['name']} → {pdf_filename}")
                else:
                    log_messages.append(f"[{timestamp}] ❌ Falló: {file_info['name']}")
                
                log_placeholder.text_area("", "\n".join(log_messages), height=200, key=f"log_done_{i}")
        
        except Exception as e:
            log_messages.append(f"[{timestamp}] ❌ Error: {file_info['name']} - {str(e)}")
            log_placeholder.text_area("", "\n".join(log_messages), height=200, key=f"log_error_{i}")
    
    # Resultado final
    status_text.empty()
    
    if converted_files:
        st.success(f"✅ Conversión completada! {len(converted_files)}/{len(all_files_to_process)} archivos convertidos")
        
        try:
            zip_path = create_zip_with_original_names(converted_files)
            
            with open(zip_path, "rb") as f:
                st.download_button(
                    label="📥 DESCARGAR PDFs CON NOMBRES ORIGINALES",
                    data=f,
                    file_name="documentos_convertidos.zip",
                    mime="application/zip",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Error creando archivo ZIP: {e}")
        
        cleanup_files([path for _, path in converted_files] + [zip_path] if 'zip_path' in locals() else [path for _, path in converted_files])
    else:
        st.error("❌ No se pudo convertir ningún archivo")

def extract_and_filter_zip(zip_path, buscar_subcarpetas):
    """Extrae archivos de un ZIP y filtra por tipos permitidos (solo .txt y .docx)"""
    allowed_extensions = ['.txt', '.docx']
    extracted_files = []
    
    with tempfile.TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Buscar archivos en el directorio extraído
        if buscar_subcarpetas:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    file_ext = Path(file).suffix.lower()
                    
                    if file_ext in allowed_extensions:
                        with open(file_path, 'rb') as f:
                            content = f.read()
                        
                        rel_path = os.path.relpath(file_path, temp_dir)
                        extracted_files.append({
                            'name': rel_path,
                            'content': content,
                            'extension': file_ext
                        })
        else:
            for item in os.listdir(temp_dir):
                item_path = os.path.join(temp_dir, item)
                if os.path.isfile(item_path):
                    file_ext = Path(item).suffix.lower()
                    
                    if file_ext in allowed_extensions:
                        with open(item_path, 'rb') as f:
                            content = f.read()
                        
                        extracted_files.append({
                            'name': item,
                            'content': content,
                            'extension': file_ext
                        })
    
    return extracted_files

def create_zip_with_original_names(converted_files):
    """Crea un archivo ZIP manteniendo los nombres originales"""
    zip_path = tempfile.NamedTemporaryFile(delete=False, suffix='.zip').name
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for original_name, file_path in converted_files:
            if os.path.exists(file_path):
                zipf.write(file_path, original_name)
    
    return zip_path

def cleanup_files(file_paths):
    """Limpia archivos temporales"""
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.unlink(file_path)
        except:
            pass

if __name__ == "__main__":
    main()
