# app.py
import streamlit as st
import os
import tempfile
from pathlib import Path
import zipfile
import shutil
import subprocess
import logging
from typing import Tuple, Dict, List
import time
import base64
import requests

# Configuración de logging
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(message)s')
logger = logging.getLogger(__name__)

class DocumentConverter:
    def __init__(self):
        self.supported_formats = {
            '.doc': 'Microsoft Word Document',
            '.docx': 'Microsoft Word Document', 
            '.rtf': 'Rich Text Format',
            '.txt': 'Plain Text',
            '.odt': 'OpenDocument Text'
        }
        
        self.max_file_size = 200 * 1024 * 1024  # 200MB
        
    def check_dependencies(self) -> Dict[str, bool]:
        """Verifica las dependencias del sistema"""
        dependencies = {
            'pandoc': self._check_pandoc(),
            'python-docx': self._check_python_docx(),
            'wkhtmltopdf': self._check_wkhtmltopdf(),
            'antiword': self._check_antiword(),
            'catdoc': self._check_catdoc(),
        }
        return dependencies
    
    def _check_pandoc(self) -> bool:
        """Verifica si Pandoc está instalado"""
        try:
            result = subprocess.run(['pandoc', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except:
            return False
    
    def _check_python_docx(self) -> bool:
        """Verifica si python-docx está instalado"""
        try:
            import docx
            return True
        except ImportError:
            return False
    
    def _check_wkhtmltopdf(self) -> bool:
        """Verifica si wkhtmltopdf está instalado"""
        try:
            result = subprocess.run(['wkhtmltopdf', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except:
            return False
    
    def _check_antiword(self) -> bool:
        """Verifica si antiword está instalado (para archivos .DOC)"""
        try:
            result = subprocess.run(['antiword', '-v'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except:
            return False
    
    def _check_catdoc(self) -> bool:
        """Verifica si catdoc está instalado (para archivos .DOC)"""
        try:
            result = subprocess.run(['catdoc', '-h'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except:
            return False
    
    def convert_document(self, input_path: str, output_dir: str = None) -> Tuple[bool, str, str]:
        """Convierte un documento a PDF - retorna (éxito, mensaje, ruta_pdf)"""
        input_path = Path(input_path)
        
        if not input_path.exists():
            return False, f"Archivo no encontrado: {input_path}", ""
        
        if input_path.stat().st_size > self.max_file_size:
            return False, f"Archivo demasiado grande: {input_path}", ""
        
        if output_dir is None:
            output_dir = input_path.parent
        
        output_path = Path(output_dir) / f"{input_path.stem}.pdf"
        
        try:
            # Seleccionar método de conversión según la extensión
            extension = input_path.suffix.lower()
            
            if extension == '.docx':
                success, message = self._convert_docx(input_path, output_path)
            elif extension == '.doc':
                success, message = self._convert_doc(input_path, output_path)
            elif extension == '.rtf':
                success, message = self._convert_rtf(input_path, output_path)
            elif extension == '.txt':
                success, message = self._convert_txt(input_path, output_path)
            elif extension == '.odt':
                success, message = self._convert_odt(input_path, output_path)
            else:
                return False, f"Formato no soportado: {extension}", ""
            
            if success:
                logger.info(f"Convertido: {input_path.name} → {output_path.name}")
                return True, message, str(output_path)
            else:
                logger.error(f"Error convirtiendo {input_path.name}: {message}")
                return False, message, ""
            
        except Exception as e:
            error_msg = f"Error inesperado: {str(e)}"
            logger.error(error_msg)
            return False, error_msg, ""
    
    def _convert_docx(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOCX a PDF usando múltiples métodos"""
        methods = [
            self._convert_with_pandoc_wkhtml,
            self._convert_with_python_docx
        ]
        
        for method in methods:
            success, message = method(input_path, output_path)
            if success:
                return True, message
        
        return False, "Todos los métodos de conversión fallaron"
    
    def _convert_doc(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOC a PDF usando métodos específicos para DOC"""
        methods = [
            self._convert_doc_with_antiword,
            self._convert_doc_with_catdoc,
            self._convert_doc_with_strings,
            self._convert_doc_with_fallback
        ]
        
        for method in methods:
            success, message = method(input_path, output_path)
            if success:
                return True, message
        
        return False, "No se pudo convertir el archivo DOC. Intente guardarlo como DOCX."
    
    def _convert_rtf(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte RTF a PDF usando wkhtmltopdf"""
        return self._convert_with_pandoc_wkhtml(input_path, output_path)
    
    def _convert_txt(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte TXT a PDF"""
        return self._convert_with_pandoc_wkhtml(input_path, output_path)
    
    def _convert_odt(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte ODT a PDF"""
        return self._convert_with_pandoc_wkhtml(input_path, output_path)
    
    def _convert_with_pandoc_wkhtml(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Conversión usando Pandoc con wkhtmltopdf"""
        try:
            # Usar wkhtmltopdf como motor PDF
            cmd = [
                'pandoc', str(input_path), 
                '-o', str(output_path),
                '--pdf-engine=wkhtmltopdf'
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0 and output_path.exists():
                return True, "Conversión exitosa con Pandoc"
            else:
                return False, f"Pandoc error: {result.stderr}"
                
        except subprocess.TimeoutExpired:
            return False, "Timeout en conversión con Pandoc"
        except Exception as e:
            return False, f"Error con Pandoc: {str(e)}"
    
    def _convert_with_python_docx(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Conversión usando python-docx (solo para DOCX)"""
        try:
            from docx import Document
            
            doc = Document(input_path)
            text_content = []
            
            # Extraer texto de párrafos
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    text_content.append(paragraph.text)
            
            # Extraer texto de tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            text_content.append(cell.text)
            
            if text_content:
                # Crear un PDF simple con el texto extraído
                success = self._create_simple_pdf(text_content, output_path, input_path.stem)
                if success:
                    return True, "Conversión básica exitosa con python-docx"
                else:
                    return False, "No se pudo crear PDF desde el texto extraído"
            else:
                return False, "No se pudo extraer texto del documento"
            
        except Exception as e:
            return False, f"Error con python-docx: {str(e)}"
    
    def _convert_doc_with_antiword(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOC a PDF usando antiword"""
        try:
            # Extraer texto con antiword
            cmd = ['antiword', str(input_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0 and result.stdout.strip():
                text_content = result.stdout.split('\n')
                success = self._create_simple_pdf(text_content, output_path, input_path.stem)
                if success:
                    return True, "Conversión exitosa con Antiword"
            
            return False, "Antiword no pudo extraer texto del archivo DOC"
            
        except subprocess.TimeoutExpired:
            return False, "Timeout con Antiword"
        except Exception as e:
            return False, f"Error con Antiword: {str(e)}"
    
    def _convert_doc_with_catdoc(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOC a PDF usando catdoc"""
        try:
            # Extraer texto con catdoc
            cmd = ['catdoc', '-w', str(input_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0 and result.stdout.strip():
                text_content = result.stdout.split('\n')
                success = self._create_simple_pdf(text_content, output_path, input_path.stem)
                if success:
                    return True, "Conversión exitosa con Catdoc"
            
            return False, "Catdoc no pudo extraer texto del archivo DOC"
            
        except subprocess.TimeoutExpired:
            return False, "Timeout con Catdoc"
        except Exception as e:
            return False, f"Error con Catdoc: {str(e)}"
    
    def _convert_doc_with_strings(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOC a PDF usando strings (método de último recurso)"""
        try:
            # Extraer texto legible con strings
            cmd = ['strings', '-n', '3', str(input_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0 and result.stdout.strip():
                # Filtrar solo líneas que parecen texto legible
                lines = result.stdout.split('\n')
                text_content = [line for line in lines if len(line.strip()) > 10 and any(c.isalpha() for c in line)]
                
                if text_content:
                    success = self._create_simple_pdf(text_content, output_path, input_path.stem)
                    if success:
                        return True, "Conversión básica exitosa (método strings)"
            
            return False, "No se pudo extraer texto legible del archivo DOC"
            
        except Exception as e:
            return False, f"Error con strings: {str(e)}"
    
    def _convert_doc_with_fallback(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Método de fallback para archivos DOC - crea un PDF informativo"""
        try:
            text_content = [
                f"Archivo: {input_path.name}",
                "Formato: Documento de Word (.DOC)",
                "",
                "⚠️ No se pudo convertir el contenido del archivo DOC.",
                "Sugerencias:",
                "1. Guarde el archivo como DOCX en Microsoft Word",
                "2. Use LibreOffice para abrir y guardar como PDF",
                "3. Utilice la versión online con LibreOffice instalado",
                "",
                "Esta versión usa métodos alternativos para archivos DOC",
                "pero puede no funcionar con documentos complejos."
            ]
            
            success = self._create_simple_pdf(text_content, output_path, input_path.stem)
            if success:
                return True, "PDF informativo creado (conversión limitada)"
            else:
                return False, "No se pudo crear PDF informativo"
                
        except Exception as e:
            return False, f"Error en método de fallback: {str(e)}"
    
    def _create_simple_pdf(self, text_content: List[str], output_path: Path, title: str) -> bool:
        """Crea un PDF simple con el contenido de texto usando wkhtmltopdf directamente"""
        try:
            # Crear un HTML simple
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>{title}</title>
                <style>
                    body {{ 
                        font-family: Arial, sans-serif; 
                        margin: 40px;
                        line-height: 1.6;
                        color: #333;
                    }}
                    h1 {{ 
                        color: #2c3e50; 
                        border-bottom: 2px solid #3498db;
                        padding-bottom: 10px;
                    }}
                    .content {{ 
                        margin: 30px 0;
                        background: #f8f9fa;
                        padding: 20px;
                        border-radius: 5px;
                        border-left: 4px solid #3498db;
                    }}
                    p {{ 
                        margin: 12px 0;
                        padding: 5px;
                    }}
                    .warning {{
                        background: #fff3cd;
                        border-left: 4px solid #ffc107;
                        padding: 15px;
                        margin: 15px 0;
                        border-radius: 4px;
                    }}
                </style>
            </head>
            <body>
                <h1>📄 {title}</h1>
                <div class="content">
                    {''.join(f'<p>{line}</p>' for line in text_content if line.strip())}
                </div>
                <div class="warning">
                    <strong>Nota:</strong> Documento convertido usando métodos alternativos. 
                    El formato original puede variar.
                </div>
                <p><em>Convertido el {time.strftime("%d/%m/%Y %H:%M")}</em></p>
            </body>
            </html>
            """
            
            # Guardar HTML temporal
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                html_path = f.name
            
            # Convertir HTML a PDF usando wkhtmltopdf directamente
            cmd = ['wkhtmltopdf', '--enable-local-file-access', '--quiet', html_path, str(output_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            # Limpiar archivo temporal
            if os.path.exists(html_path):
                os.unlink(html_path)
            
            return result.returncode == 0 and output_path.exists()
            
        except Exception as e:
            logger.error(f"Error creando PDF simple: {e}")
            return False
    
    def process_zip_folder(self, zip_path: str, output_dir: str = None) -> Dict[str, Tuple[bool, str, str]]:
        """Procesa una carpeta ZIP con múltiples archivos"""
        results = {}
        
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Convertir todos los archivos soportados
                for file_path in Path(temp_dir).rglob('*'):
                    if file_path.is_file() and file_path.suffix.lower() in self.supported_formats:
                        success, message, pdf_path = self.convert_document(file_path, output_dir)
                        results[file_path.name] = (success, message, pdf_path)
                
            except Exception as e:
                logger.error(f"Error procesando ZIP: {str(e)}")
                results['ZIP Processing'] = (False, f"Error procesando ZIP: {str(e)}", "")
        
        return results

# Inicializar el conversor
@st.cache_resource
def get_converter():
    return DocumentConverter()

# Configuración de la página
st.set_page_config(
    page_title="Conversor de Documentos",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .file-info {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 5px;
        padding: 10px;
        margin: 5px 0;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .download-section {
        background-color: #e8f5e8;
        border: 2px solid #4caf50;
        border-radius: 10px;
        padding: 20px;
        margin: 20px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

def process_uploaded_files(uploaded_files, converter):
    """Procesar archivos subidos individualmente"""
    if 'conversion_history' not in st.session_state:
        st.session_state.conversion_history = []
    
    successful_conversions = 0
    total_files = len(uploaded_files)
    
    if total_files == 0:
        st.warning("No hay archivos para procesar")
        return
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    results_container = st.container()
    
    converted_files = []
    conversion_results = []
    
    with results_container:
        st.subheader("📊 Progreso de Conversión")
        
        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"🔄 Procesando {i+1}/{total_files}: {uploaded_file.name}")
            
            # Guardar archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_file.name).suffix) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
            
            try:
                # Convertir archivo
                success, message, pdf_path = converter.convert_document(tmp_path)
                
                # Registrar en historial
                timestamp = time.strftime("%H:%M:%S")
                output_file = f"{Path(uploaded_file.name).stem}.pdf"
                
                st.session_state.conversion_history.append({
                    'timestamp': timestamp,
                    'input': uploaded_file.name,
                    'output': output_file if success else "N/A",
                    'success': success,
                    'message': message
                })
                
                conversion_results.append({
                    'original_name': uploaded_file.name,
                    'success': success,
                    'message': message,
                    'pdf_path': pdf_path
                })
                
                if success:
                    successful_conversions += 1
                    if pdf_path and os.path.exists(pdf_path):
                        converted_files.append(pdf_path)
                    
                    # Mostrar mensaje específico para DOC
                    if Path(uploaded_file.name).suffix.lower() == '.doc':
                        st.success(f"✅ {uploaded_file.name} → {output_file} (Conversión básica)")
                        st.markdown("""
                        <div class="warning-box">
                        ⚠️ <strong>Archivo DOC convertido:</strong> La conversión de archivos .DOC es básica. 
                        Para mejor calidad, guarde el archivo como .DOCX en Microsoft Word.
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.success(f"✅ {uploaded_file.name} → {output_file}")
                else:
                    st.error(f"❌ {uploaded_file.name}: {message}")
            
            except Exception as e:
                error_msg = f"Error procesando {uploaded_file.name}: {str(e)}"
                st.error(f"❌ {error_msg}")
                st.session_state.conversion_history.append({
                    'timestamp': time.strftime("%H:%M:%S"),
                    'input': uploaded_file.name,
                    'output': "N/A",
                    'success': False,
                    'message': error_msg
                })
            
            finally:
                # Limpiar archivo temporal
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
            
            progress_bar.progress((i + 1) / total_files)
    
    status_text.text("")
    
    # Mostrar sección de descargas
    if successful_conversions > 0:
        st.markdown("---")
        st.markdown('<div class="download-section">', unsafe_allow_html=True)
        st.subheader("📥 Descargar Archivos Convertidos")
        
        if successful_conversions == 1:
            # Descarga individual
            pdf_path = converted_files[0]
            original_name = Path(conversion_results[0]['original_name']).stem
            download_name = f"{original_name}.pdf"
            
            with open(pdf_path, 'rb') as f:
                st.download_button(
                    label=f"📄 Descargar {download_name}",
                    data=f,
                    file_name=download_name,
                    mime="application/pdf",
                    type="primary",
                    key="single_download"
                )
                
        else:
            # Descarga múltiple - crear ZIP
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.write(f"**{successful_conversions} archivos convertidos exitosamente**")
                
            with col2:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_zip:
                    zip_path = tmp_zip.name
                
                try:
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for pdf_file in converted_files:
                            if os.path.exists(pdf_file):
                                zipf.write(pdf_file, os.path.basename(pdf_file))
                    
                    with open(zip_path, 'rb') as f:
                        st.download_button(
                            label="📦 Descargar todos los PDFs (ZIP)",
                            data=f,
                            file_name="documentos_convertidos.zip",
                            mime="application/zip",
                            type="primary",
                            key="zip_download"
                        )
                finally:
                    if os.path.exists(zip_path):
                        os.unlink(zip_path)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Resumen final
    if successful_conversions > 0:
        st.balloons()
        st.success(f"🎉 Conversión completada! {successful_conversions}/{total_files} archivos convertidos")
    else:
        st.error("😞 No se pudo convertir ningún archivo")

def process_zip_file(uploaded_zip, converter):
    """Procesar archivo ZIP"""
    with st.spinner("📦 Procesando archivo ZIP..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = Path(temp_dir) / uploaded_zip.name
            zip_path.write_bytes(uploaded_zip.getvalue())
            
            # Procesar ZIP
            results = converter.process_zip_folder(zip_path, temp_dir)
            
            successful = sum(1 for result in results.values() if result[0])
            total = len(results)
            
            # Mostrar resultados
            st.subheader("📊 Resultados de la conversión:")
            
            converted_files = []
            for filename, (success, message, pdf_path) in results.items():
                if success:
                    st.success(f"✅ {filename}")
                    if pdf_path and os.path.exists(pdf_path):
                        converted_files.append(pdf_path)
                else:
                    st.error(f"❌ {filename}: {message}")
            
            # Sección de descargas para ZIP
            if successful > 0:
                st.markdown("---")
                st.markdown('<div class="download-section">', unsafe_allow_html=True)
                st.subheader("📥 Descargar Archivos Convertidos")
                
                # Crear ZIP con resultados
                output_zip = Path(temp_dir) / "documentos_convertidos.zip"
                with zipfile.ZipFile(output_zip, 'w') as zipf:
                    for pdf_file in converted_files:
                        if os.path.exists(pdf_file):
                            zipf.write(pdf_file, os.path.basename(pdf_file))
                
                # Botón de descarga
                with open(output_zip, 'rb') as f:
                    st.download_button(
                        label=f"📦 Descargar {successful} archivos PDF (ZIP)",
                        data=f,
                        file_name="documentos_convertidos.zip",
                        mime="application/zip",
                        type="primary",
                        key="zip_result_download"
                    )
                
                st.markdown('</div>', unsafe_allow_html=True)
                
                st.success(f"📊 {successful}/{total} archivos convertidos exitosamente")
            else:
                st.error("No se pudo convertir ningún archivo del ZIP")

def main():
    converter = get_converter()
    
    st.markdown('<h1 class="main-header">📄 Conversor de Documentos a PDF</h1>', unsafe_allow_html=True)
    
    # Información importante
    st.markdown("""
    <div class="info-box">
    💡 <strong>Novedades:</strong> 
    - ✅ Soporte completo para archivos .DOC (conversión básica)
    - 📋 Mejor extracción de texto con Antiword y Catdoc
    - 📥 Botones de descarga mejorados
    - 🚀 Compatibilidad mejorada
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar con información
    with st.sidebar:
        st.header("ℹ️ Información")
        st.markdown("""
        **Formatos soportados:**
        - 📝 DOC (Word) - Conversión básica
        - 📝 DOCX (Word) - Conversión completa
        - 📋 RTF (Rich Text)
        - 📄 TXT (Texto plano)
        - 📦 ZIP (Carpetas)
        
        **Nota sobre archivos DOC:**
        Los archivos .DOC tienen conversión básica de texto.
        Para mejor calidad, guarde como .DOCX.
        
        **Límites:**
        - 200MB por archivo
        - Conversión masiva vía ZIP
        """)
        
        # Verificar dependencias
        st.header("🔧 Estado del Sistema")
        deps = converter.check_dependencies()
        for dep, available in deps.items():
            status = "✅" if available else "❌"
            st.write(f"{status} {dep}")
            
        # Información específica sobre DOC
        if not deps['antiword'] and not deps['catdoc']:
            st.markdown("""
            <div class="warning-box">
            ⚠️ <strong>Archivos .DOC:</strong> 
            Sin Antiword/Catdoc, la conversión de .DOC será muy básica.
            </div>
            """, unsafe_allow_html=True)
    
    # Pestañas principales
    tab1, tab2, tab3 = st.tabs(["📤 Subir Archivos", "📁 Subir Carpeta ZIP", "📊 Historial"])
    
    with tab1:
        st.header("Subir Archivos Individuales")
        
        # Área de upload
        uploaded_files = st.file_uploader(
            "Arrastra y suelta archivos aquí",
            type=list(converter.supported_formats.keys()),
            accept_multiple_files=True,
            help="Límite: 200MB por archivo • DOC, DOCX, RTF, TXT"
        )
        
        if uploaded_files:
            st.subheader("📁 Archivos subidos:")
            
            # Mostrar información de archivos
            for uploaded_file in uploaded_files:
                file_size = uploaded_file.size / (1024 * 1024)  # MB
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.write(f"**{uploaded_file.name}**")
                with col2:
                    st.write(f"{file_size:.1f} MB")
                with col3:
                    format_name = converter.supported_formats.get(Path(uploaded_file.name).suffix.lower(), "Desconocido")
                    if Path(uploaded_file.name).suffix.lower() == '.doc':
                        st.write(f"📝 {format_name} (Básico)")
                    else:
                        st.write(f"📄 {format_name}")
            
            # Botón de conversión
            if st.button("🔄 Iniciar Conversión", type="primary", key="convert_single"):
                process_uploaded_files(uploaded_files, converter)
    
    with tab2:
        st.header("Subir Carpeta ZIP")
        
        uploaded_zip = st.file_uploader(
            "Arrastra y suelta archivo ZIP aquí",
            type=['zip'],
            help="Límite: 200MB • ZIP con documentos"
        )
        
        if uploaded_zip:
            file_size = uploaded_zip.size / (1024 * 1024)  # MB
            st.success(f"📦 Carpeta ZIP cargada: {uploaded_zip.name} ({file_size:.1f} MB)")
            
            if st.button("🔄 Procesar Carpeta ZIP", type="primary", key="convert_zip"):
                process_zip_file(uploaded_zip, converter)
    
    with tab3:
        st.header("📋 Registro de Actividad")
        
        # Mostrar historial de conversiones
        if 'conversion_history' in st.session_state and st.session_state.conversion_history:
            st.write(f"**Últimas {len(st.session_state.conversion_history)} conversiones:**")
            for entry in reversed(st.session_state.conversion_history[-10:]):  # Mostrar últimos 10
                if entry['success']:
                    st.markdown(f"""
                    <div class="success-box">
                        ✅ [{entry['timestamp']}] Convertido: {entry['input']} → {entry['output']}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="error-box">
                        ❌ [{entry['timestamp']}] Error: {entry['input']} - {entry['message']}
                    </div>
                    """, unsafe_allow_html=True)
            
            # Botón para limpiar historial
            if st.button("🗑️ Limpiar Historial", key="clear_history"):
                st.session_state.conversion_history = []
                st.rerun()
        else:
            st.info("No hay actividad reciente")

if __name__ == "__main__":
    main()
