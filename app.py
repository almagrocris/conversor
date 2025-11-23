import streamlit as st
import os
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime
import subprocess
import shutil
import sys

# Configuración de la página
st.set_page_config(
    page_title="Conversor PDF Web - Mejorado",
    page_icon="🔄",
    layout="wide"
)

def main():
    st.title("🔄 CONVERSOR PDF WEB - MEJORADO")
    st.markdown("**@Cristobal Almagro**")
    st.markdown("---")
    
    # Sidebar con configuración
    with st.sidebar:
        st.header("⚙️ Configuración")
        
        st.subheader("Formatos a convertir:")
        doc = st.checkbox("📄 Word (.doc, .docx)", value=True)
        txt = st.checkbox("📝 Texto (.txt)", value=True)
        rtf = st.checkbox("📋 Texto enriquecido (.rtf)", value=True)
        odt = st.checkbox("📓 OpenDocument (.odt)", value=True)
        
        st.subheader("📁 Opciones de entrada:")
        subir_zip = st.checkbox("📦 Permitir subir carpetas (ZIP)", value=True)
        buscar_subcarpetas = st.checkbox("🔍 Buscar en subcarpetas", value=True)
        
        # Botón para verificar LibreOffice
        if st.button("🔍 Verificar LibreOffice"):
            check_libreoffice()
        
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
        if doc: allowed_types.extend(['doc', 'docx'])
        if txt: allowed_types.append('txt')
        if rtf: allowed_types.append('rtf')
        if odt: allowed_types.append('odt')
        
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
        **Nuevas características:**
        - ✅ **Nombres originales preservados** - Los PDFs mantendrán el nombre del archivo original
        - ✅ **Soporte para carpetas** - Sube archivos ZIP con múltiples documentos
        - ✅ **Búsqueda recursiva** - Busca archivos en subcarpetas dentro de ZIPs
        - ✅ **Botón SALIR** - Cierra la aplicación fácilmente
        
        **Formatos soportados:**
        - 📄 .doc, .docx (Word) - via LibreOffice
        - 📝 .txt (Texto) - via ReportLab  
        - 📋 .rtf (Texto enriquecido) - via LibreOffice
        - 📓 .odt (OpenDocument) - via LibreOffice
        """)

def check_libreoffice():
    """Verifica si LibreOffice está instalado y accesible"""
    # En Streamlit Cloud, mostrar mensaje especial
    if is_running_on_streamlit():
        st.info("🌐 **Modo Streamlit Cloud Activado**")
        st.warning("⚠️ LibreOffice no está disponible en Streamlit Cloud")
        st.info("💡 Funcionalidades disponibles: Conversión de archivos TXT")
        return
    
    libreoffice_paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice.bin",
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice"
    ]
    
    found = False
    for path in libreoffice_paths:
        if os.path.exists(path):
            st.success(f"✅ LibreOffice encontrado: {path}")
            
            # Probar la versión
            try:
                result = subprocess.run([path, '--version'], capture_output=True, text=True, timeout=10)
                if result.returncode == 0:
                    st.info(f"📋 Versión: {result.stdout.strip()}")
                found = True
            except:
                st.warning(f"⚠️ No se pudo verificar la versión en: {path}")
    
    if not found:
        st.error("❌ LibreOffice no encontrado. Instálalo desde: https://www.libreoffice.org/")

def is_running_on_streamlit():
    """Detecta si la aplicación está ejecutándose en Streamlit Cloud"""
    return 'STREAMLIT_SHARING' in os.environ or 'STREAMLIT_SERVER' in os.environ

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
    allowed_extensions = ['.doc', '.docx', '.txt', '.rtf', '.odt']
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
    """Convierte un archivo a PDF usando el nombre especificado"""
    try:
        # Guardar archivo temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file_info['name']).suffix) as tmp_input:
            tmp_input.write(file_info['content'])
            input_path = tmp_input.name
        
        # Archivo de salida con nombre específico
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        
        extension = file_info['extension']
        
        if extension == '.txt':
            success = convert_txt(input_path, output_path)
        else:
            # Para Office documents, usar LibreOffice
            success = convert_with_libreoffice(input_path, output_path)
        
        # Limpiar archivo temporal de entrada
        if os.path.exists(input_path):
            os.unlink(input_path)
        
        return output_path if success else None
        
    except Exception as e:
        st.error(f"Error en conversión de {file_info['name']}: {e}")
        return None

def convert_with_libreoffice(input_path, output_path):
    """Convierte archivos de Office a PDF usando LibreOffice"""
    try:
        # Rutas posibles de LibreOffice en macOS
        libreoffice_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice.bin",
            "/opt/homebrew/bin/soffice",
            "/usr/local/bin/soffice"
        ]
        
        libreoffice_cmd = None
        
        # Buscar LibreOffice
        for path in libreoffice_paths:
            if os.path.exists(path):
                libreoffice_cmd = path
                break
        
        if not libreoffice_cmd:
            st.error("❌ LibreOffice no encontrado. Verifica que esté instalado en /Applications/")
            return False
        
        # Crear directorio temporal para la conversión
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Ejecutar LibreOffice en modo headless
            result = subprocess.run([
                libreoffice_cmd,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', temp_dir,
                input_path
            ], capture_output=True, text=True, timeout=60)
            
            # Verificar si la conversión fue exitosa
            if result.returncode != 0:
                st.error(f"❌ Error en LibreOffice: {result.stderr}")
                return False
            
            # Buscar el PDF generado
            pdf_files = list(Path(temp_dir).glob("*.pdf"))
            
            if not pdf_files:
                st.error("❌ LibreOffice no generó ningún archivo PDF")
                return False
            
            # Mover el PDF a la ubicación final con el nombre deseado
            generated_pdf = pdf_files[0]
            shutil.move(str(generated_pdf), output_path)
            
            return True
            
        except subprocess.TimeoutExpired:
            st.error("❌ Timeout: LibreOffice tardó demasiado en convertir el archivo")
            return False
        except Exception as e:
            st.error(f"❌ Error durante la conversión: {e}")
            return False
        finally:
            # Limpiar directorio temporal
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
                
    except Exception as e:
        st.error(f"❌ Error general en LibreOffice: {e}")
        return False

def convert_txt(input_path, output_path):
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
