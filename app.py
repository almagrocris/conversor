# app.py
import streamlit as st
import os
import tempfile
from pathlib import Path
import zipfile
import shutil
from document_converter import DocumentConverter
import time

# Configuración de la página
st.set_page_config(
    page_title="Conversor de Documentos",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicializar el conversor
@st.cache_resource
def get_converter():
    return DocumentConverter()

converter = get_converter()

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
</style>
""", unsafe_allow_html=True)

def main():
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
    
    # Pestañas principales
    tab1, tab2, tab3 = st.tabs(["📤 Subir Archivos", "📁 Subir Carpeta ZIP", "📊 Progreso"])
    
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
                process_uploaded_files(uploaded_files)
    
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
                process_zip_file(uploaded_zip)
    
    with tab3:
        st.header("Registro de Actividad")
        
        # Mostrar historial de conversiones
        if 'conversion_history' in st.session_state:
            for entry in st.session_state.conversion_history:
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

def process_uploaded_files(uploaded_files):
    """Procesar archivos subidos individualmente"""
    if 'conversion_history' not in st.session_state:
        st.session_state.conversion_history = []
    
    successful_conversions = 0
    total_files = len(uploaded_files)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
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
                st.success(f"✅ {uploaded_file.name} → {output_file}")
            else:
                st.error(f"❌ {uploaded_file.name}: {message}")
        
        except Exception as e:
            st.error(f"❌ Error procesando {uploaded_file.name}: {str(e)}")
        
        finally:
            # Limpiar archivo temporal
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
        
        progress_bar.progress((i + 1) / total_files)
    
    status_text.text("")
    
    # Resumen final
    if successful_conversions > 0:
        st.balloons()
        st.success(f"🎉 Conversión completada! {successful_conversions}/{total_files} archivos convertidos")
        
        # Botón de descarga (podrías implementar la creación de un ZIP con todos los PDFs)
        if st.button("📥 Descargar PDFs", key="download_pdfs"):
            st.info("🔧 Función de descarga masiva en desarrollo...")
    else:
        st.error("😞 No se pudo convertir ningún archivo")

def process_zip_file(uploaded_zip):
    """Procesar archivo ZIP"""
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
                    mime="application/zip"
                )
        else:
            st.error("No se pudo convertir ningún archivo del ZIP")

if __name__ == "__main__":
    main()
