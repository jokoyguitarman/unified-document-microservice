from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import tempfile
import base64
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont
import io

# Document processing libraries
from pptx import Presentation
from docx import Document
import pandas as pd
from pdf2image import convert_from_bytes
import fitz  # PyMuPDF for PDF processing

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with your frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def create_text_image(text_content: str, page_number: int = 1, title: str = "Document") -> str:
    """Create a visual representation of text content as an image"""
    # Create a blank image
    width, height = 1200, 800
    img = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(img)
    
    # Add title and page number
    draw.text((50, 30), f"{title} - Page {page_number}", fill='black')
    
    # Add text content with word wrapping
    y_position = 80
    words = text_content.split()
    current_line = ""
    
    for word in words:
        test_line = current_line + " " + word if current_line else word
        # Simple word wrapping (you can enhance this with proper text measurement)
        if len(test_line) > 80:  # Approximate characters per line
            if current_line:
                draw.text((50, y_position), current_line, fill='black')
                y_position += 20
                current_line = word
            else:
                draw.text((50, y_position), word, fill='black')
                y_position += 20
        else:
            current_line = test_line
    
    # Draw remaining text
    if current_line:
        draw.text((50, y_position), current_line, fill='black')
    
    # Convert to base64
    buffer = BytesIO()
    img.save(buffer, format='PNG')
    img_base64 = base64.b64encode(buffer.getvalue()).decode()
    return img_base64

def process_word_document(file_path: str) -> list[str]:
    """Convert Word document to images"""
    doc = Document(file_path)
    images = []
    
    # Extract text from paragraphs
    text_content = ""
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            text_content += paragraph.text + "\n\n"
    
    # Split content into pages (simple approach - you can enhance this)
    words_per_page = 500
    words = text_content.split()
    pages = []
    
    for i in range(0, len(words), words_per_page):
        page_words = words[i:i + words_per_page]
        pages.append(" ".join(page_words))
    
    # Create images for each page
    for i, page_content in enumerate(pages):
        if page_content.strip():
            image = create_text_image(page_content, i + 1, "Word Document")
            images.append(image)
    
    return images

def process_excel_document(file_path: str) -> list[str]:
    """Convert Excel document to images"""
    # Read Excel file
    excel_file = pd.ExcelFile(file_path)
    images = []
    
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Convert DataFrame to text representation
        text_content = f"Sheet: {sheet_name}\n\n"
        text_content += df.to_string(index=False)
        
        image = create_text_image(text_content, 1, f"Excel - {sheet_name}")
        images.append(image)
    
    return images

def process_powerpoint_document(file_path: str) -> list[str]:
    """Convert PowerPoint document to images"""
    prs = Presentation(file_path)
    images = []
    
    for i, slide in enumerate(prs.slides):
        # Extract text from shapes
        text_content = f"Slide {i + 1}\n\n"
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text_content += shape.text + "\n"
        
        image = create_text_image(text_content, i + 1, "PowerPoint")
        images.append(image)
    
    return images

def process_pdf_document(file_path: str) -> list[str]:
    """Convert PDF document to images"""
    # Read PDF and convert to images
    with open(file_path, 'rb') as file:
        pdf_bytes = file.read()
    
    images_pil = convert_from_bytes(pdf_bytes)
    images_base64 = []
    
    for img in images_pil:
        # Convert PIL image to base64
        buffer = BytesIO()
        img.save(buffer, format='PNG')
        img_base64 = base64.b64encode(buffer.getvalue()).decode()
        images_base64.append(img_base64)
    
    return images_base64

def process_text_document(file_path: str) -> list[str]:
    """Convert text document to images"""
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Split content into pages
    words_per_page = 500
    words = content.split()
    pages = []
    
    for i in range(0, len(words), words_per_page):
        page_words = words[i:i + words_per_page]
        pages.append(" ".join(page_words))
    
    images = []
    for i, page_content in enumerate(pages):
        if page_content.strip():
            image = create_text_image(page_content, i + 1, "Text Document")
            images.append(image)
    
    return images

@app.post("/document-to-images")
async def document_to_images(file: UploadFile = File(...)):
    # Add detailed logging
    print(f"Received file upload request:")
    print(f"  Filename: {file.filename}")
    print(f"  Content type: {file.content_type}")
    print(f"  File size: {file.size if hasattr(file, 'size') else 'unknown'}")
    
    # Validate file type
    filename = file.filename.lower() if file.filename else ""
    supported_extensions = ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt', '.txt']
    
    print(f"  File extension check: {filename}")
    print(f"  Supported extensions: {supported_extensions}")
    
    if not filename:
        print("  ERROR: No filename provided")
        raise HTTPException(
            status_code=400, 
            detail="No filename provided"
        )
    
    if not any(filename.endswith(ext) for ext in supported_extensions):
        print(f"  ERROR: Unsupported file type: {filename}")
        raise HTTPException(
            status_code=400, 
            detail=f"Unsupported file type: {filename}. Supported: {', '.join(supported_extensions)}"
        )
    
    print(f"  File validation passed, processing...")
    
    # Save uploaded file to a temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name
    
    try:
        # Process based on file type
        if filename.endswith('.pdf'):
            print("  Processing as PDF...")
            images = process_pdf_document(tmp_path)
        elif filename.endswith(('.docx', '.doc')):
            print("  Processing as Word document...")
            images = process_word_document(tmp_path)
        elif filename.endswith(('.xlsx', '.xls')):
            print("  Processing as Excel document...")
            images = process_excel_document(tmp_path)
        elif filename.endswith(('.pptx', '.ppt')):
            print("  Processing as PowerPoint document...")
            images = process_powerpoint_document(tmp_path)
        elif filename.endswith('.txt'):
            print("  Processing as text document...")
            images = process_text_document(tmp_path)
        else:
            print(f"  ERROR: Unsupported file type after validation: {filename}")
            raise HTTPException(status_code=400, detail="Unsupported file type")
        
        print(f"  Processing completed successfully. Generated {len(images)} images.")
        
        return {
            "images": images, 
            "num_pages": len(images),
            "file_type": os.path.splitext(filename)[1][1:].upper()
        }
        
    except Exception as e:
        print(f"  ERROR during processing: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error processing document: {str(e)}")
    finally:
        os.remove(tmp_path)

@app.post("/image-to-images")
async def image_to_images(file: UploadFile = File(...)):
    """Handle single image upload"""
    # Validate that the file is an image
    supported_image_types = ['image/jpeg', 'image/jpg', 'image/png', 'image/gif', 'image/webp', 'image/bmp']
    
    if file.content_type not in supported_image_types:
        raise HTTPException(
            status_code=400, 
            detail=f"Unsupported image type: {file.content_type}. Supported: {', '.join(supported_image_types)}"
        )
    
    try:
        # Read image file
        image_bytes = await file.read()
        
        # Convert to PIL Image for processing
        img = Image.open(BytesIO(image_bytes))
        
        # Convert to base64
        buffer = BytesIO()
        img.save(buffer, format='PNG')
        img_base64 = base64.b64encode(buffer.getvalue()).decode()
        
        return {
            "images": [img_base64],
            "num_pages": 1,
            "file_type": "IMAGE",
            "message": f"Processed 1 image"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing image: {str(e)}")

@app.post("/images-to-images")
async def images_to_images(files: list[UploadFile] = File(...)):
    """Handle multiple image uploads as a document sequence"""
    if not files:
        raise HTTPException(status_code=400, detail="No files provided")
    
    # Handle single file case (convert to list)
    if not isinstance(files, list):
        files = [files]
    
    # Validate that all files are images
    supported_image_types = ['image/jpeg', 'image/jpg', 'image/png', 'image/gif', 'image/webp', 'image/bmp']
    
    for file in files:
        if file.content_type not in supported_image_types:
            raise HTTPException(
                status_code=400, 
                detail=f"Unsupported image type: {file.content_type}. Supported: {', '.join(supported_image_types)}"
            )
    
    images_base64 = []
    
    try:
        for i, file in enumerate(files):
            # Read image file
            image_bytes = await file.read()
            
            # Convert to PIL Image for processing
            img = Image.open(BytesIO(image_bytes))
            
            # Convert to base64
            buffer = BytesIO()
            img.save(buffer, format='PNG')
            img_base64 = base64.b64encode(buffer.getvalue()).decode()
            
            images_base64.append(img_base64)
        
        return {
            "images": images_base64,
            "num_pages": len(images_base64),
            "file_type": "IMAGE_SEQUENCE",
            "message": f"Processed {len(images_base64)} images as document sequence"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing images: {str(e)}")

@app.get("/")
def root():
    return {
        "message": "Unified Document to Image Microservice",
        "supported_formats": ["PDF", "DOCX", "DOC", "XLSX", "XLS", "PPTX", "PPT", "TXT"],
        "endpoints": {
            "document_to_images": "/document-to-images",
            "images_to_images": "/images-to-images"
        },
        "features": [
            "Single document upload (PDF, Word, Excel, PowerPoint, Text)",
            "Multiple image uploads as document sequence",
            "Returns base64 images ready for AI analysis"
        ]
    }

@app.get("/health")
def health_check():
    return {
        "status": "healthy",
        "service": "Unified Document to Image Microservice",
        "timestamp": "2024-01-01T00:00:00Z"
    } 