import streamlit as st
import os
from pdf2image import convert_from_path
from PIL import Image
import time
import tempfile
import base64
from datetime import datetime
import io
import requests
import json
from docx2pdf import convert
import pythoncom
from ollama import Client

# Ollama configuration
MODEL = "qwen2.5vl:7b"
SERVER = "http://43.205.68.27:11434/"
# SERVER = "http://13.203.132.235:11434/"

# Initialize Ollama client
client = Client(host=SERVER)

st.set_page_config(
    page_title="OCR Text Extraction",
    page_icon="🖼️",
    layout="wide"
)

# Custom CSS for clean UI (unchanged)
st.markdown("""
<style>
    .main-title {
        font-size: 32px !important;
        font-weight: 700 !important;
        color: #1f3a60 !important;
        margin-bottom: 10px !important;
        text-align: center;
    }
    
    .main-sub {
        font-size: 18px !important;
        color: #2d3748 !important;
        text-align: center;
        margin-bottom: 30px !important;
    }
    
    .section-header {
        font-size: 22px !important;
        font-weight: 600 !important;
        color: #1f3a60 !important;
        border-bottom: 2px solid #4a90e2;
        padding-bottom: 8px;
        margin-bottom: 15px;
    }
    
    .extract-button {
        background: linear-gradient(135deg, #6a11cb 0%, #2575fc 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 18px !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 28px !important;
        transition: all 0.3s ease !important;
        width: 100%;
        margin: 20px 0;
    }
    
    .extract-button:hover {
        background: linear-gradient(135deg, #2575fc 0%, #6a11cb 100%) !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    .info-box {
        background-color: #f8fafc;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 15px;
        border-left: 4px solid #4a90e2;
    }
    
    .metric-value {
        font-size: 20px !important;
        font-weight: 700 !important;
        color: #1f3a60 !important;
        margin: 5px 0;
    }
    
    .metric-label {
        font-size: 14px !important;
        color: #718096 !important;
        margin: 0;
    }
    
    .docx-preview {
        font-family: Arial, sans-serif;
        line-height: 1.5;
        color: #212529;
        max-height: 370px;
        overflow: auto;
        padding: 5px;
    }
    
    .docx-preview h1, .docx-preview h2, .docx-preview h3, .docx-preview h4 {
        color: #1f3a60;
        margin-top: 1em;
        margin-bottom: 0.5em;
    }
    
    .docx-preview p {
        margin-bottom: 1em;
    }
    
    .docx-preview ul, .docx-preview ol {
        margin-left: 1.5em;
        margin-bottom: 1em;
    }
    
    .pdf-viewer {
        width: 100%;
        height: 400px;
        border: none;
        border-radius: 4px;
    }
    
    .stButton button {
        width: 100%;
    }
    
    /* Fix alignment issues */
    .column-container {
        display: flex;
        flex-direction: column;
        height: 100%;
    }
    
    .upload-section {
        flex: 1;
        display: flex;
        flex-direction: column;
    }
    
    .button-section {
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        height: 100%;
        padding: 20px 0;
    }
    
    .result-section {
        flex: 1;
        display: flex;
        flex-direction: column;
    }
</style>
""", unsafe_allow_html=True)

def extract_text_with_ollama(image_path: str) -> str:
    prompt = (
        "You are an OCR assistant. Extract ALL visible text from this image exactly as it appears. "
        "Include:\n"
        "- All printed text\n"
        "- All handwritten text\n"
        "- Numbers, dates, amounts\n"
        "- Text in any language\n"
        "- Preserve line breaks and formatting\n\n"
        "Return ONLY the extracted text content without any commentary or explanations."
    )
    
    try:
        response = client.chat(
            model=MODEL,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                    "images": [image_path]
                }
            ]
        )
        return response["message"]["content"].strip()
    except Exception as e:
        raise Exception(f"Ollama API error: {str(e)}")

def process_pdf(pdf_path: str) -> tuple:
    try:
        pages = convert_from_path(pdf_path, dpi=200)
    except Exception as e:
        raise Exception(f"Failed to convert PDF to images: {str(e)}")

    all_text = []
    total_time = 0
    
    for i, page in enumerate(pages, start=1):
        temp_file = None
        temp_file_path = None
        try:
            # Create temporary file for the page
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                temp_file_path = temp_file.name
                page.save(temp_file_path, "PNG")
            
            # Extract text from the page
            start_time = time.time()
            page_text = extract_text_with_ollama(temp_file_path)
            end_time = time.time()
            
            elapsed = end_time - start_time
            total_time += elapsed
            
            # Add page text with timing info ---> LP Removed the timings and page info on the extracted data..
            # all_text.append(f"\n--- Page {i} ---\n{page_text}\n\n(Time taken: {elapsed:.2f} seconds)")
            all_text.append(f"{page_text}")
            
        except Exception as e:
            raise Exception(f"Error processing page {i}: {str(e)}")
        finally:
            # Clean up temporary file
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.unlink(temp_file_path)
                except:
                    pass

    return "\n".join(all_text), total_time

def process_image(image_path: str) -> tuple:
    start_time = time.time()
    extracted_text = extract_text_with_ollama(image_path)
    end_time = time.time()
    elapsed = end_time - start_time
    return extracted_text, elapsed

def process_docx(docx_path: str) -> tuple:
    """Process DOCX file by converting to PDF first, then using Qwen model"""
    try:
        # Create temporary directory
        temp_dir = tempfile.mkdtemp()
        
        # Convert DOCX to PDF
        pdf_path = os.path.join(temp_dir, "converted.pdf")
        pythoncom.CoInitialize()
        convert(docx_path, pdf_path)
        pythoncom.CoUninitialize()
        
        # Process the converted PDF with Qwen model
        extracted_text, processing_time = process_pdf(pdf_path)
        
        # Clean up
        try:
            os.remove(pdf_path)
            os.rmdir(temp_dir)
        except:
            pass
            
        return extracted_text, processing_time
        
    except Exception as e:
        # Fallback to text extraction if conversion fails
        try:
            import docx
            start_time = time.time()
            doc = docx.Document(docx_path)
            full_text = []
            for paragraph in doc.paragraphs:
                full_text.append(paragraph.text)
            end_time = time.time()
            elapsed = end_time - start_time
            return '\n'.join(full_text), elapsed
        except ImportError:
            raise Exception("python-docx library is required for DOCX processing. Install with: pip install python-docx")
        except Exception as fallback_error:
            raise Exception(f"Failed to process DOCX file: {str(e)}. Fallback also failed: {str(fallback_error)}")

def analyze_accuracy(extracted_text: str) -> float:
    # Simple heuristic-based accuracy estimation
    if not extracted_text or len(extracted_text.strip()) == 0:
        return 0.0
    
    # Count lines and words
    lines = extracted_text.split('\n')
    non_empty_lines = [line for line in lines if line.strip()]
    word_count = sum(len(line.split()) for line in non_empty_lines)
    
    # Simple heuristic: more content = higher confidence
    if word_count > 100:
        confidence = 0.85 + (min(word_count, 1000) / 10000)  # Cap at 0.95 for very long documents
    elif word_count > 50:
        confidence = 0.75 + (word_count / 500)
    elif word_count > 10:
        confidence = 0.6 + (word_count / 100)
    else:
        confidence = 0.3 + (word_count / 33)  # 0.3 to 0.6 for very short texts
    
    return min(confidence, 0.95)  # Cap at 95%

def format_time_display(seconds):
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    else:
        minutes = int(seconds // 60)
        remaining_seconds = seconds % 60
        return f"{minutes} min {remaining_seconds:.2f} sec"

def display_pdf(file):
    # Read file as bytes
    bytes_data = file.getvalue()
    # Encode to base64
    base64_pdf = base64.b64encode(bytes_data).decode('utf-8')
    # Create PDF display HTML
    pdf_display = f'''
    <iframe 
        class="pdf-viewer"
        src="data:application/pdf;base64,{base64_pdf}#toolbar=0&navpanes=0&scrollbar=1"
        frameborder="0"
        scrolling="yes"
    ></iframe>
    '''
    return pdf_display

def display_image(file):
    # Display image using PIL
    image = Image.open(file)
    st.image(image, use_container_width=True)
    
def display_docx(file):
    """Convert DOCX to PDF and display it with the same PDF viewer style"""
    try:
        # Create a temporary directory
        temp_dir = tempfile.mkdtemp()
        
        # Save DOCX temporarily
        docx_path = os.path.join(temp_dir, file.name)
        with open(docx_path, "wb") as f:
            f.write(file.getvalue())

        # Convert DOCX to PDF with COM init
        pdf_path = os.path.join(temp_dir, "converted.pdf")
        pythoncom.CoInitialize()
        convert(docx_path, pdf_path)
        pythoncom.CoUninitialize()

        # Read the converted PDF and display it with the same style as PDF files
        with open(pdf_path, "rb") as f:
            bytes_data = f.read()
            base64_pdf = base64.b64encode(bytes_data).decode('utf-8')
            
        # Use the same PDF viewer style as display_pdf()
        pdf_display = f'''
        <iframe 
            class="pdf-viewer"
            src="data:application/pdf;base64,{base64_pdf}#toolbar=0&navpanes=0&scrollbar=1"
            frameborder="0"
            scrolling="yes"
        ></iframe>
        '''
            
        # Clean up
        try:
            os.remove(docx_path)
            os.remove(pdf_path)
            os.rmdir(temp_dir)
        except:
            pass
            
        return pdf_display
    except Exception as e:
        return f"<div class='docx-preview'>Error previewing DOCX file: {str(e)}</div>"

def main():
    st.markdown('<p class="main-title">🖼️ OCR Text Extraction</p>', unsafe_allow_html=True)
    st.markdown('<p class="main-sub">Extract text from documents and images using Ollama with Qwen2.5-VL</p>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'extracted_text' not in st.session_state:
        st.session_state.extracted_text = ""
    if 'processing_time' not in st.session_state:
        st.session_state.processing_time = 0
    if 'confidence' not in st.session_state:
        st.session_state.confidence = 0
    
    # Create three columns layout
    col1, col2, col3 = st.columns([2, 1, 2], gap="large")
    
    with col1:
        st.markdown('<div class="column-container">', unsafe_allow_html=True)
        st.markdown('<div class="section-header">📁 Upload & Preview</div>', unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Upload a document or image",
            type=['jpg', 'jpeg', 'png', 'pdf', 'docx'],
            help="Supported formats: JPG, PNG, PDF, DOCX"
        )
        
        if uploaded_file:
            st.markdown('<div class="preview-container">', unsafe_allow_html=True)
            
            file_type = uploaded_file.type
            if file_type == "application/pdf":
                pdf_display = display_pdf(uploaded_file)
                st.markdown(pdf_display, unsafe_allow_html=True)
            elif file_type in ["image/jpeg", "image/jpg", 'image/png']:
                display_image(uploaded_file)
            elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                docx_preview = display_docx(uploaded_file)
                st.markdown(docx_preview, unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="preview-container center-container">', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="button-section">', unsafe_allow_html=True)
        
        if uploaded_file:
            if st.button("Extract Text", key="extract_btn", use_container_width=True):
                with st.spinner("Processing..."):
                    temp_file = None
                    temp_file_path = None
                    try:
                        # Save uploaded file to temporary location
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}")
                        temp_file.write(uploaded_file.getvalue())
                        temp_file_path = temp_file.name
                        temp_file.close()
                        
                        # Process based on file type
                        file_type = uploaded_file.type
                        if file_type == "application/pdf":
                            extracted_text, processing_time = process_pdf(temp_file_path)
                        elif file_type in ["image/jpeg", "image/jpg", 'image/png']:
                            extracted_text, processing_time = process_image(temp_file_path)
                        elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                            extracted_text, processing_time = process_docx(temp_file_path)
                        else:
                            st.error("Unsupported file type")
                            return
                        
                        # Analyze accuracy
                        confidence = analyze_accuracy(extracted_text)
                        
                        # Store results in session state
                        st.session_state.extracted_text = extracted_text
                        st.session_state.processing_time = processing_time
                        st.session_state.confidence = confidence
                        
                    except Exception as e:
                        st.error(f"Error processing file: {str(e)}")
                    finally:
                        if temp_file_path and os.path.exists(temp_file_path):
                            try:
                                os.unlink(temp_file_path)
                            except:
                                pass
        else:
            st.info("Upload a file to enable extraction")
            
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="column-container">', unsafe_allow_html=True)
        st.markdown('<div class="section-header">📝 Extracted Text</div>', unsafe_allow_html=True)
        
        if st.session_state.extracted_text:
            # Display metrics
            col_time, col_conf = st.columns(2)
            with col_time:
                st.markdown('<div class="metric-box">', unsafe_allow_html=True)
                st.markdown('<p class="metric-label">Time Taken</p>', unsafe_allow_html=True)
                st.markdown(f'<p class="metric-value">{format_time_display(st.session_state.processing_time)}</p>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_conf:
                st.markdown('<div class="metric-box">', unsafe_allow_html=True)
                st.markdown('<p class="metric-label">Confidence</p>', unsafe_allow_html=True)
                st.markdown(f'<p class="metric-value">{st.session_state.confidence * 100:.1f}%</p>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
            
            # Display extracted text
            st.markdown('<div class="text-container">', unsafe_allow_html=True)
            st.text_area(
                "Extracted Text:",
                value=st.session_state.extracted_text,
                height=300,
                label_visibility="collapsed"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Download button
            st.download_button(
                label="Download Extracted Text",
                data=st.session_state.extracted_text,
                file_name=f"{uploaded_file.name.split('.')[0]}_extracted.txt",
                mime="text/plain",
                use_container_width=True
            )
        else:
            st.markdown('<div class="text-container center-container">', unsafe_allow_html=True)
            st.info("Extracted text will appear here after processing")
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()