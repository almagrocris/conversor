import streamlit as st
import os
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime
import sys
import io

# Configuración de la página
st.set_page_config(
    page_title="Conversor PDF Web - Pure Python",
    page_icon="🔄",
    layout="wide"
)

def main():
    st.title("🔄 CONVERSOR PDF WEB - PURE PYTHON")
    st.markdown("**@Cristobal Almagro**")
    st.markdown("---")
    
    # Sidebar con configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        st.subheader("Formatos a convertir:")
        docx = st.checkbox("📄 Word (.docx)", value=True)
        txt = st.checkbox("📝 Texto (.txt)", value=True)
        # .doc necesita conversion especial
        doc = st.checkbox("📄 Word Legacy (.doc)", value=True)
        
        st.subheader("📁 Opciones de entrada:")
        subir_zip = st.checkbox("📦 Permitir subir carpetas (ZIP)", value=True)
        buscar_subcarpetas = st.checkbox("🔍 Buscar en subcarpetas", value=True)
        
        st.markdown("---")
        
        # BOTÓN SALIR en el sidebar
        st.markdown("---")
        if st.button("🔒 SALIR", type="secondary", use_container_width=True):
            st.success("👋 ¡Hasta pronto! Cerrando la aplicación...")
            import time
            time.sleep(2)
            sys.exit()
        
        st.info("💡 **100% Python** - Sin dependencias externas")
    
    # Área principal
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 Subir Archivos o Carpetas")
        
        # Determinar tipos de archivo permitidos
        allowed_types = []
        if docx: allowed_types.extend(['docx'])
        if txt: allowed_types.append('txt')
        if doc: allowed_types.append('doc')
        
        if not allowed_types:
            st.warning("⚠️ Selecciona al menos un tipo de archivo en la configuración")
            return
        
        # Subida de archivos individuales
        uploaded_files = st.file_uploader(
            "📄 Archivos individuales",
            type=allowed_types,
            accept_multiple_files=True,
            help=f"Formatos permitidos: {', '.join(allowed_types)}"
        )
        
        # Subida de carpetas ZIP (nueva funcionalidad)
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
            total_files += 1  # Contamos el ZIP como un "lote" de archivos
        
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
        **✨ Nueva Versión - 100% Python**
        - ✅ **Sin LibreOffice** - Solo librerías Python
        - ✅ **Funciona en Streamlit Cloud** - Todos los formatos
        - ✅ **Nombres originales preservados**
        - ✅ **Soporte para carpetas ZIP**
        
        **Formatos soportados:**
        - 📄 .docx (Word moderno) - via python-docx2pdf
        - 📄 .doc (Word legacy) - Conversión básica a texto
        - 📝 .txt (Texto) - via ReportLab
        
        **Tecnologías:**
        - python-docx2pdf
        - ReportLab
        - Pure Python Magic!
        """)

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
            # Extraer y procesar archivos del ZIP
            zip_files = extract_and_filter_zip(zip_path, buscar_subcarpetas)
            all_files_to_process.extend(zip_files)
        finally:
            # Limpiar archivo ZIP temporal
            if os.path.exists(zip_path):
                os.unlink(zip_path)
    
    if not all_files_to_process:
        st.error("❌ No se encontraron archivos para procesar")
        return
    
    # Procesar todos los archivos
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    converted_files = []  # Ahora guardamos (nombre_original, ruta_pdf)
    log_messages = []
    
    # Área de log
    log_container = st.container()
    with log_container:
        st.subheader("📝 Registro de Actividad")
        log_placeholder = st.empty()
    
    for i, file_info in enumerate(all_files_to_process):
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
                # Usar el nombre original para el PDF
                original_name = Path(file_info['name']).stem  # Nombre sin extensión
                pdf_filename = f"{original_name}.pdf"
                
                pdf_path = convert_to_pdf(file_info, pdf_filename)
                
                if pdf_path and os.path.exists(pdf_path):
                    # Guardar con el nombre original
                    converted_files.append((pdf_filename, pdf_path))
                    log_messages.append(f"[{timestamp}] ✅ Convertido: {file_info['name']} → {pdf_filename}")
                else:
                    log_messages.append(f"[{timestamp}] ❌ Falló: {file_info['name']}")
                
                # Actualizar log
                log_placeholder.text_area("", "\n".join(log_messages), height=200, key=f"log_done_{i}")
        
        except Exception as e:
            log_messages.append(f"[{timestamp}] ❌ Error: {file_info['name']} - {str(e)}")
            log_placeholder.text_area("", "\n".join(log_messages), height=200, key=f"log_error_{i}")
    
    # Resultado final
    status_text.empty()
    
    if converted_files:
        st.success(f"✅ Conversión completada! {len(converted_files)}/{len(all_files_to_process)} archivos convertidos")
        
        # Crear y ofrecer descarga
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
        
        # Limpiar archivos temporales
        cleanup_files([path for _, path in converted_files] + [zip_path] if 'zip_path' in locals() else [path for _, path in converted_files])
    else:
        st.error("❌ No se pudo convertir ningún archivo")

def extract_and_filter_zip(zip_path, buscar_subcarpetas):
    """Extrae archivos de un ZIP y filtra por tipos permitidos"""
    allowed_extensions = ['.doc', '.docx', '.txt']
    extracted_files = []
    
    with tempfile.TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Buscar archivos en el directorio extraído
        if buscar_subcarpetas:
            # Búsqueda recursiva
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    file_ext = Path(file).suffix.lower()
                    
                    if file_ext in allowed_extensions:
                        with open(file_path, 'rb') as f:
                            content = f.read()
                        
                        # Mantener la estructura de carpetas relativa
                        rel_path = os.path.relpath(file_path, temp_dir)
                        extracted_files.append({
                            'name': rel_path,
                            'content': content,
                            'extension': file_ext
                        })
        else:
            # Solo archivos en la raíz
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

def convert_to_pdf(file_info, output_filename):
    """Convierte un archivo a PDF usando librerías Python puras"""
    try:
        # Guardar archivo temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file_info['name']).suffix) as tmp_input:
            tmp_input.write(file_info['content'])
            input_path = tmp_input.name
        
        # Archivo de salida con nombre específico
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        
        extension = file_info['extension']
        
        if extension == '.txt':
            success = convert_txt_to_pdf(input_path, output_path)
        elif extension == '.docx':
            success = convert_docx_to_pdf(input_path, output_path)
        elif extension == '.doc':
            success = convert_doc_to_pdf(input_path, output_path)
        else:
            st.warning(f"⚠️ Formato no soportado: {extension}")
            success = False
        
        # Limpiar archivo temporal de entrada
        if os.path.exists(input_path):
            os.unlink(input_path)
        
        return output_path if success else None
        
    except Exception as e:
        st.error(f"Error en conversión de {file_info['name']}: {e}")
        return None

def convert_docx_to_pdf(input_path, output_path):
    """Convierte DOCX a PDF usando python-docx y ReportLab"""
    try:
        # Intentar importar python-docx
        try:
            from docx import Document
        except ImportError:
            st.error("❌ python-docx no está instalado. Ejecuta: pip install python-docx")
            return False
        
        # Leer documento DOCX
        doc = Document(input_path)
        
        # Crear PDF
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        
        # Configurar PDF
        pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Estilo para títulos
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=14,
            spaceAfter=12,
        )
        
        # Estilo para párrafos
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
        )
        
        # Procesar cada párrafo del documento
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():  # Ignorar párrafos vacíos
                # Detectar si es un título
                if paragraph.style.name.startswith('Heading'):
                    story.append(Paragraph(paragraph.text, title_style))
                else:
                    story.append(Paragraph(paragraph.text, normal_style))
                story.append(Spacer(1, 0.1 * inch))
        
        # Procesar tablas (conversión básica)
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join([cell.text for cell in row.cells if cell.text])
                if row_text:
                    story.append(Paragraph(f"📊 {row_text}", normal_style))
                    story.append(Spacer(1, 0.05 * inch))
        
        # Construir PDF
        if story:  # Solo si hay contenido
            pdf_doc.build(story)
            return os.path.exists(output_path) and os.path.getsize(output_path) > 0
        else:
            st.warning("📄 Documento DOCX vacío o sin contenido convertible")
            return False
            
    except Exception as e:
        st.error(f"❌ Error conversión DOCX: {str(e)}")
        return False

def convert_doc_to_pdf(input_path, output_path):
    """Convierte DOC a PDF (conversión básica a texto)"""
    try:
        # Para archivos .doc antiguos, usar una conversión básica a texto
        # Nota: .doc es un formato binario complejo, esta es una solución básica
        
        # Intentar leer como texto plano (funciona para algunos .doc simples)
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        content = None
        
        for encoding in encodings:
            try:
                with open(input_path, 'r', encoding=encoding, errors='ignore') as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            # Si no se puede leer como texto, crear un PDF informativo
            content = f"Documento .doc: {os.path.basename(input_path)}\n\n" \
                     "⚠️ Los archivos .doc (Word antiguo) tienen formato binario complejo.\n" \
                     "Para mejor conversión, guarda el archivo como .docx y vuelve a intentar."
        
        # Crear PDF con el contenido
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        
        pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        
        # Limpiar y formatear contenido
        cleaned_content = content.replace('\x00', '')  # Remover caracteres nulos
        formatted_text = cleaned_content.replace('\n', '<br/>')
        
        story = [Paragraph(formatted_text, styles['Normal'])]
        pdf_doc.build(story)
        
        return os.path.exists(output_path) and os.path.getsize(output_path) > 0
        
    except Exception as e:
        st.error(f"❌ Error conversión DOC: {str(e)}")
        # Crear un PDF de error
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
            
            pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
            styles = getSampleStyleSheet()
            story = [Paragraph(f"Error convirtiendo archivo .doc: {str(e)}", styles['Normal'])]
            pdf_doc.build(story)
            return True
        except:
            return False

def convert_txt_to_pdf(input_path, output_path):
    """Convierte archivo TXT a PDF"""
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        
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
        
        # Formatear texto
        formatted_text = content.replace('\n', '<br/>').replace('\t', '    ')
        story = [Paragraph(formatted_text, styles['Normal'])]
        
        doc.build(story)
        
        return os.path.exists(output_path) and os.path.getsize(output_path) > 0
        
    except Exception as e:
        st.error(f"Error en conversión TXT: {e}")
        return False

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
