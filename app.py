# app.py
import streamlit as st
import os
import tempfile
from pathlib import Path
import zipfile
import shutil
import subprocess
import logging
from typing import Tuple, Dict
import time

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
            'libreoffice': self._check_libreoffice(),
            'pandoc': self._check_pandoc(),
        }
        return dependencies
    
    def _check_libreoffice(self) -> bool:
        """Verifica si LibreOffice está instalado"""
        try:
            result = subprocess.run(['soffice', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except:
            return False
    
    def _check_pandoc(self) -> bool:
        """Verifica si Pandoc está instalado"""
        try:
            result = subprocess.run(['pandoc', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            return result.returncode == 0
        except:
            return False
    
    def convert_document(self, input_path: str, output_dir: str = None) -> Tuple[bool, str]:
        """Convierte un documento a PDF"""
        input_path = Path(input_path)
        
        if not input_path.exists():
            return False, f"Archivo no encontrado: {input_path}"
        
        if input_path.stat().st_size > self.max_file_size:
            return False, f"Archivo demasiado grande: {input_path}"
        
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
                return False, f"Formato no soportado: {extension}"
            
            if success:
                logger.info(f"Convertido: {input_path.name} → {output_path.name}")
            else:
                logger.error(f"Error convirtiendo {input_path.name}: {message}")
            
            return success, message
            
        except Exception as e:
            error_msg = f"Error inesperado: {str(e)}"
            logger.error(error_msg)
            return False, error_msg
    
    def _convert_docx(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOCX a PDF"""
        return self._convert_with_libreoffice(input_path, output_path)
    
    def _convert_doc(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte DOC a PDF"""
        return self._convert_with_libreoffice(input_path, output_path)
    
    def _convert_rtf(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte RTF a PDF"""
        return self._convert_with_libreoffice(input_path, output_path)
    
    def _convert_txt(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte TXT a PDF"""
        return self._convert_with_libreoffice(input_path, output_path)
    
    def _convert_odt(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Convierte ODT a PDF"""
        return self._convert_with_libreoffice(input_path, output_path)
    
    def _convert_with_libreoffice(self, input_path: Path, output_path: Path) -> Tuple[bool, str]:
        """Conversión usando LibreOffice (método más robusto)"""
        try:
            cmd = [
                'soffice', '--headless', '--convert-to', 'pdf',
                '--outdir', str(output_path.parent),
                str(input_path)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0:
                # Verificar si el archivo PDF fue creado
                expected_path = output_path.parent / f"{input_path.stem}.pdf"
                if expected_path.exists():
                    return True, "Conversión exitosa con LibreOffice"
            
            return False, f"LibreOffice error: {result.stderr}"
            
        except subprocess.TimeoutExpired:
            return False, "Timeout en conversión con LibreOffice"
        except Exception as e:
            return False, f"Error con LibreOffice: {str(e)}"
    
    def process_zip_folder(self, zip_path: str, output_dir: str = None) -> Dict[str, Tuple[bool, str]]:
        """Procesa una carpeta ZIP con múltiples archivos"""
        results = {}
        
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Convertir todos los archivos soportados
                for file_path in Path(temp_dir).rglob('*'):
                    if file_path.is_file() and file_path.suffix.lower() in self.supported_formats:
                        success, message = self.convert_document(file_path, output_dir)
                        results[file_path.name] = (success, message)
                
            except Exception as e:
                logger.error(f"Error procesando ZIP: {str(e)}")
                results['ZIP Processing'] = (False, f"Error procesando ZIP: {str(e)}")
        
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
    
    with results_container:
        st.subheader("Resultados de la conversión:")
        
        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Procesando {i+1}/{total_files}: {uploaded_file.name}")
            
            # Guardar archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded_file.name).suffix) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
            
            try:
                # Convertir archivo
                success, message = converter.convert_document(tmp_path)
                
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
                
                if success:
                    successful_conversions += 1
                    # Guardar ruta del archivo convertido
                    pdf_path = Path(tmp_path).parent / output_file
                    if pdf_path.exists():
                        converted_files.append(pdf_path)
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
    
    # Resumen final y descargas
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if successful_conversions > 0:
            st.balloons()
            st.success(f"🎉 Conversión completada! {successful_conversions}/{total_files} archivos convertidos")
        else:
            st.error("😞 No se pudo convertir ningún archivo")
    
    with col2:
        # Botones de descarga
        if successful_conversions == 1 and converted_files:
            pdf_path = converted_files[0]
            with open(pdf_path, 'rb') as f:
                st.download_button(
                    label="📥 Descargar PDF",
                    data=f,
                    file_name=pdf_path.name,
                    mime="application/pdf",
                    type="primary"
                )
        elif successful_conversions > 1:
            # Crear ZIP con múltiples PDFs
            with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_zip:
                zip_path = tmp_zip.name
            
            try:
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for pdf_file in converted_files:
                        if pdf_file.exists():
                            zipf.write(pdf_file, pdf_file.name)
                
                with open(zip_path, 'rb') as f:
                    st.download_button(
                        label="📥 Descargar todos los PDFs (ZIP)",
                        data=f,
                        file_name="documentos_convertidos.zip",
                        mime="application/zip",
                        type="primary"
                    )
            finally:
                if os.path.exists(zip_path):
                    os.unlink(zip_path)

def process_zip_file(uploaded_zip, converter):
    """Procesar archivo ZIP"""
    with st.spinner("Procesando archivo ZIP..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = Path(temp_dir) / uploaded_zip.name
            zip_path.write_bytes(uploaded_zip.getvalue())
            
            # Procesar ZIP
            results = converter.process_zip_folder(zip_path, temp_dir)
            
            successful = sum(1 for result in results.values() if result[0])
            total = len(results)
            
            # Mostrar resultados
            st.subheader("Resultados de la conversión:")
            
            for filename, (success, message) in results.items():
                if success:
                    st.success(f"✅ {filename}")
                else:
                    st.error(f"❌ {filename}: {message}")
            
            if successful > 0:
                st.success(f"📊 {successful}/{total} archivos convertidos exitosamente")
                
                # Crear ZIP con resultados
                output_zip = Path(temp_dir) / "converted_pdfs.zip"
                with zipfile.ZipFile(output_zip, 'w') as zipf:
                    for pdf_file in Path(temp_dir).glob("*.pdf"):
                        zipf.write(pdf_file, pdf_file.name)
                
                # Botón de descarga
                with open(output_zip, 'rb') as f:
                    st.download_button(
                        label="📥 Descargar PDFs en ZIP",
                        data=f,
                        file_name="documentos_convertidos.zip",
                        mime="application/zip",
                        type="primary"
                    )
            else:
                st.error("No se pudo convertir ningún archivo del ZIP")

def main():
    converter = get_converter()
    
    st.markdown('<h1 class="main-header">📄 Conversor de Documentos a PDF</h1>', unsafe_allow_html=True)
    
    # Sidebar con información
    with st.sidebar:
        st.header("ℹ️ Información")
        st.markdown("""
        **Formatos soportados:**
        - 📝 DOC, DOCX (Word)
        - 📋 RTF (Rich Text)
        - 📄 TXT (Texto plano)
        - 📦 ZIP (Carpetas)
        
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
            
        if not any(deps.values()):
            st.error("Se requiere al menos LibreOffice o Pandoc para la conversión")
    
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
            st.subheader("Archivos subidos:")
            
            # Mostrar información de archivos
            for uploaded_file in uploaded_files:
                file_size = uploaded_file.size / (1024 * 1024)  # MB
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.write(f"**{uploaded_file.name}**")
                with col2:
                    st.write(f"{file_size:.1f} MB")
                with col3:
                    st.write(converter.supported_formats.get(Path(uploaded_file.name).suffix.lower(), "Desconocido"))
            
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
            st.success(f"📦 Carpeta ZIP cargada: {uploaded_zip.name}")
            
            if st.button("🔄 Procesar Carpeta ZIP", type="primary", key="convert_zip"):
                process_zip_file(uploaded_zip, converter)
    
    with tab3:
        st.header("Registro de Actividad")
        
        # Mostrar historial de conversiones
        if 'conversion_history' in st.session_state and st.session_state.conversion_history:
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
        else:
            st.info("No hay actividad reciente")

if __name__ == "__main__":
    main()
