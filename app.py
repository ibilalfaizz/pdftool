import streamlit as st
import os
import tempfile
from pathlib import Path
import io
from PIL import Image

# PowerPoint to PDF imports
try:
    from pptx import Presentation
    from reportlab.lib.pagesizes import A4
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# PDF to TIFF imports
try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

# Page configuration
st.set_page_config(
    page_title="File Format Converter",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #667eea;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 0.5rem 1rem;
    }
    .stButton>button:hover {
        background-color: #764ba2;
    }
    .info-box {
        padding: 1rem;
        background-color: #f0f2f6;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .footer {
        text-align: center;
        padding: 2rem 0;
        color: #666;
        margin-top: 3rem;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="main-header">
        <h1>🔄 File Format Converter</h1>
        <p>Convert PowerPoint to PDF and PDF to TIFF with ease</p>
    </div>
""", unsafe_allow_html=True)

# Initialize session state
if 'last_dpi' not in st.session_state:
    st.session_state.last_dpi = 300
if 'last_compression' not in st.session_state:
    st.session_state.last_compression = 'tiff_deflate'

def convert_pptx_to_pdf(uploaded_file):
    """Convert PowerPoint file to PDF"""
    try:
        # Reset file pointer
        uploaded_file.seek(0)
        
        # Check file extension
        file_ext = Path(uploaded_file.name).suffix.lower()
        if file_ext == '.ppt':
            st.warning("⚠️ Note: .ppt files (older format) may not work perfectly. For best results, use .pptx files.")
        
        # Save uploaded file to temporary location
        suffix = '.pptx' if file_ext in ['.ppt', '.pptx'] else '.pptx'
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_path = tmp_file.name
        
        # Load presentation (python-pptx only supports .pptx)
        if file_ext == '.ppt':
            raise ValueError("The .ppt format (PowerPoint 97-2003) is not directly supported. Please convert your file to .pptx format first, or use LibreOffice to convert it.")
        
        prs = Presentation(tmp_path)
        
        # Create PDF
        pdf_buffer = io.BytesIO()
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        
        doc = SimpleDocTemplate(pdf_buffer, pagesize=A4,
                               rightMargin=72, leftMargin=72,
                               topMargin=72, bottomMargin=18)
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor='#333333',
            spaceAfter=30,
            alignment=TA_CENTER
        )
        body_style = ParagraphStyle(
            'CustomBody',
            parent=styles['Normal'],
            fontSize=11,
            textColor='#000000',
            spaceAfter=12,
            alignment=TA_LEFT
        )
        
        story = []
        
        # Convert each slide
        for i, slide in enumerate(prs.slides):
            if i > 0:
                story.append(PageBreak())
            
            # Add slide title
            story.append(Paragraph(f"Slide {i + 1}", title_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Extract and add text content
            text_content = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    text_content.append(shape.text.strip())
            
            if text_content:
                for text in text_content:
                    # Clean and format text
                    clean_text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    # Split into paragraphs
                    paragraphs = clean_text.split('\n')
                    for para in paragraphs:
                        if para.strip():
                            story.append(Paragraph(para.strip(), body_style))
                            story.append(Spacer(1, 0.1*inch))
            else:
                story.append(Paragraph("(No text content on this slide)", body_style))
        
        # Build PDF
        doc.build(story)
        pdf_buffer.seek(0)
        
        # Clean up
        os.unlink(tmp_path)
        
        return pdf_buffer.getvalue()
    
    except Exception as e:
        st.error(f"Error converting PowerPoint to PDF: {str(e)}")
        import traceback
        st.error(f"Details: {traceback.format_exc()}")
        return None

def convert_pdf_to_tiff(pdf_bytes, dpi=300, compression='tiff_deflate'):
    """Convert PDF to multi-page TIFF"""
    try:
        # Convert PDF to images
        images = convert_from_bytes(pdf_bytes, dpi=dpi)
        
        if not images:
            return None
        
        # Save as multi-page TIFF
        tiff_buffer = io.BytesIO()
        
        # Determine compression
        if compression == 'tiff_deflate':
            comp = 'tiff_deflate'
        elif compression == 'tiff_lzw':
            comp = 'tiff_lzw'
        elif compression == 'tiff_adobe_deflate':
            comp = 'tiff_adobe_deflate'
        else:
            comp = None
        
        # Save first image to get format
        if len(images) == 1:
            images[0].save(tiff_buffer, format='TIFF', compression=comp)
        else:
            images[0].save(
                tiff_buffer,
                format='TIFF',
                compression=comp,
                save_all=True,
                append_images=images[1:]
            )
        
        tiff_buffer.seek(0)
        return tiff_buffer.getvalue()
    
    except Exception as e:
        st.error(f"Error converting PDF to TIFF: {str(e)}")
        return None

def get_pdf_preview(pdf_bytes):
    """Get first page of PDF as image for preview"""
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150, first_page=1, last_page=1)
        if images:
            return images[0]
        return None
    except Exception as e:
        st.warning(f"Could not generate PDF preview: {str(e)}")
        return None

def get_tiff_preview(tiff_bytes):
    """Get first page of TIFF as image for preview"""
    try:
        img = Image.open(io.BytesIO(tiff_bytes))
        return img
    except Exception as e:
        st.warning(f"Could not generate TIFF preview: {str(e)}")
        return None

# Sidebar for navigation
st.sidebar.title("Navigation")
tool_choice = st.sidebar.radio(
    "Select Tool",
    ["PowerPoint to PDF", "PDF to TIFF"],
    index=0
)

# Main content area
if tool_choice == "PowerPoint to PDF":
    st.header("📄 PowerPoint to PDF Converter")
    
    if not PPTX_AVAILABLE:
        st.error("⚠️ Required libraries not installed. Please install python-pptx and reportlab.")
        st.code("pip install python-pptx reportlab", language="bash")
    else:
        st.markdown("""
        <div class="info-box">
            <strong>📋 Instructions:</strong><br>
            1. Upload a .pptx file (recommended) or .ppt file<br>
            2. Click "Convert to PDF"<br>
            3. Preview and download your PDF<br>
            <small>Note: .pptx files work best. For .ppt files, convert to .pptx first for optimal results.</small>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Choose a PowerPoint file",
            type=['ppt', 'pptx'],
            help="Upload a .ppt or .pptx file to convert to PDF"
        )
        
            if uploaded_file is not None:
            # Show file info
            uploaded_file.seek(0)
            file_size = len(uploaded_file.read())
            uploaded_file.seek(0)
            st.info(f"📁 File: {uploaded_file.name} ({file_size / 1024:.2f} KB)")
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🔄 Convert to PDF", type="primary"):
                    with st.spinner("Converting PowerPoint to PDF... This may take a moment."):
                        uploaded_file.seek(0)
                        pdf_bytes = convert_pptx_to_pdf(uploaded_file)
                        
                        if pdf_bytes:
                            st.session_state['pdf_output'] = pdf_bytes
                            st.session_state['pdf_filename'] = Path(uploaded_file.name).stem + ".pdf"
                            output_size = len(pdf_bytes)
                            st.success(f"✅ Conversion successful! Output size: {output_size / 1024:.2f} KB")
            
            if 'pdf_output' in st.session_state:
                with col2:
                    st.download_button(
                        label="📥 Download PDF",
                        data=st.session_state['pdf_output'],
                        file_name=st.session_state['pdf_filename'],
                        mime="application/pdf",
                        type="primary"
                    )
                
                # Preview
                st.subheader("📊 Preview")
                preview_img = get_pdf_preview(st.session_state['pdf_output'])
                if preview_img:
                    st.image(preview_img, caption="First page preview", use_container_width=True)
                else:
                    st.info("Preview not available for this PDF")

elif tool_choice == "PDF to TIFF":
    st.header("🖼️ PDF to TIFF Converter")
    
    if not PDF2IMAGE_AVAILABLE:
        st.error("⚠️ Required libraries not installed. Please install pdf2image and Pillow.")
        st.code("pip install pdf2image pillow", language="bash")
        st.info("""
        **Note:** pdf2image requires Poppler. Install it:
        - **macOS:** `brew install poppler`
        - **Linux:** `sudo apt-get install poppler-utils`
        - **Windows:** Download from [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)
        """)
    else:
        st.markdown("""
        <div class="info-box">
            <strong>📋 Instructions:</strong><br>
            1. Upload a .pdf file<br>
            2. Select DPI (higher = better quality, larger file)<br>
            3. Choose compression method<br>
            4. Click "Convert to TIFF"<br>
            5. Preview and download your TIFF
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Choose a PDF file",
            type=['pdf'],
            help="Upload a .pdf file to convert to multi-page TIFF"
        )
        
        if uploaded_file is not None:
            # Show file info
            uploaded_file.seek(0)
            file_size = len(uploaded_file.read())
            uploaded_file.seek(0)
            st.info(f"📁 File: {uploaded_file.name} ({file_size / 1024:.2f} KB)")
            
            # Options
            col1, col2 = st.columns(2)
            
            with col1:
                dpi = st.selectbox(
                    "DPI (Resolution)",
                    [150, 300, 600],
                    index=[150, 300, 600].index(st.session_state.last_dpi) if st.session_state.last_dpi in [150, 300, 600] else 1,
                    help="Higher DPI = better quality but larger file size. 300 DPI is recommended for most uses."
                )
                st.session_state.last_dpi = dpi
            
            with col2:
                compression = st.selectbox(
                    "Compression Method",
                    ['tiff_deflate', 'tiff_lzw', 'tiff_adobe_deflate', 'None'],
                    index=['tiff_deflate', 'tiff_lzw', 'tiff_adobe_deflate', 'None'].index(st.session_state.last_compression) if st.session_state.last_compression in ['tiff_deflate', 'tiff_lzw', 'tiff_adobe_deflate', 'None'] else 0,
                    help="Compression reduces file size. Deflate is recommended for best balance."
                )
                st.session_state.last_compression = compression
            
            # Convert button
            if st.button("🔄 Convert to TIFF", type="primary"):
                with st.spinner("Converting PDF to TIFF... This may take a moment."):
                    pdf_bytes = uploaded_file.read()
                    tiff_bytes = convert_pdf_to_tiff(
                        pdf_bytes,
                        dpi=dpi,
                        compression=compression if compression != 'None' else None
                    )
                    
                    if tiff_bytes:
                        st.session_state['tiff_output'] = tiff_bytes
                        st.session_state['tiff_filename'] = Path(uploaded_file.name).stem + ".tiff"
                        st.success("✅ Conversion successful!")
            
            # Download button and preview
            if 'tiff_output' in st.session_state:
                output_size = len(st.session_state['tiff_output'])
                st.info(f"📁 Output file size: {output_size / 1024:.2f} KB")
                
                st.download_button(
                    label="📥 Download TIFF",
                    data=st.session_state['tiff_output'],
                    file_name=st.session_state['tiff_filename'],
                    mime="image/tiff",
                    type="primary"
                )
                
                # Preview
                st.subheader("📊 Preview")
                preview_img = get_tiff_preview(st.session_state['tiff_output'])
                if preview_img:
                    st.image(preview_img, caption="First page preview", use_container_width=True)
                else:
                    st.info("Preview not available for this TIFF")

# Footer
st.markdown("""
    <div class="footer">
        <hr>
        <p>© 2024 File Format Converter | Built with Streamlit</p>
        <p style="font-size: 0.8em;">Supports: PowerPoint (PPT/PPTX) → PDF | PDF → TIFF</p>
    </div>
""", unsafe_allow_html=True)

