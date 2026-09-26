from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import subprocess
import os
import uuid
import shutil
from pdf2docx import Converter
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches

app = FastAPI(title="Basic SNU Document Converter")

# Allow requests from your Vercel frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # Change to your vercel URL in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "/tmp/conversions"
os.makedirs(TEMP_DIR, exist_ok=True)

def cleanup_session(session_dir: str):
    """
    CRITICAL PRIVACY GUARANTEE:
    This function physically deletes the temporary folder and all its contents
    from the server memory/disk immediately after the file is sent back to the user.
    """
    if os.path.exists(session_dir):
        shutil.rmtree(session_dir, ignore_errors=True)


@app.get("/")
def health_check():
    return {"status": "healthy", "service": "basic-snu-converter", "privacy": "stateless"}


@app.post("/convert/to-pdf")
async def convert_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Converts Word (.docx) and PowerPoint (.pptx) to PDF using LibreOffice"""
    if not file.filename.lower().endswith(('.doc', '.docx', '.ppt', '.pptx')):
        raise HTTPException(status_code=400, detail="Invalid file type.")
        
    session_id = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)
    
    # Schedule absolute deletion of the file after the request completes
    background_tasks.add_task(cleanup_session, session_dir)
    
    input_path = os.path.join(session_dir, file.filename)
    
    try:
        # Save file to RAM/tmp
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Run LibreOffice to convert
        subprocess.run(
            ["libreoffice", "--headless", "--nologo", "--nofirststartwizard", 
             "--convert-to", "pdf", input_path, "--outdir", session_dir],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True
        )
        
        base_name = os.path.splitext(file.filename)[0]
        output_pdf_path = os.path.join(session_dir, f"{base_name}.pdf")
        
        if not os.path.exists(output_pdf_path):
            raise HTTPException(status_code=500, detail="Conversion failed.")
            
        return FileResponse(path=output_pdf_path, filename=f"{base_name}.pdf", media_type="application/pdf")
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/convert/pdf-to-word")
async def convert_pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Converts PDF to Word using pdf2docx"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Must upload a PDF file.")
        
    session_id = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)
    
    background_tasks.add_task(cleanup_session, session_dir)
    
    input_path = os.path.join(session_dir, file.filename)
    base_name = os.path.splitext(file.filename)[0]
    output_docx_path = os.path.join(session_dir, f"{base_name}.docx")
    
    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Extract text and layout to word
        cv = Converter(input_path)
        cv.convert(output_docx_path)
        cv.close()
        
        if not os.path.exists(output_docx_path):
            raise HTTPException(status_code=500, detail="Conversion to Word failed.")
            
        return FileResponse(
            path=output_docx_path, 
            filename=f"{base_name}.docx", 
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/convert/pdf-to-ppt")
async def convert_pdf_to_ppt(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """Converts PDF to PPT by rendering pages as slide images (since true PDF->PPT parsing is proprietary)"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Must upload a PDF file.")
        
    session_id = str(uuid.uuid4())
    session_dir = os.path.join(TEMP_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)
    
    background_tasks.add_task(cleanup_session, session_dir)
    
    input_path = os.path.join(session_dir, file.filename)
    base_name = os.path.splitext(file.filename)[0]
    output_pptx_path = os.path.join(session_dir, f"{base_name}.pptx")
    
    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Create presentation
        prs = Presentation()
        # Set slide size to standard 4:3 (can be adjusted to 16:9)
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        blank_slide_layout = prs.slide_layouts[6]
        
        # Render PDF pages to images
        pdf_document = fitz.open(input_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap(dpi=150)
            img_path = os.path.join(session_dir, f"page_{page_num}.png")
            pix.save(img_path)
            
            # Add image to slide
            slide = prs.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
        prs.save(output_pptx_path)
        pdf_document.close()
        
        return FileResponse(
            path=output_pptx_path, 
            filename=f"{base_name}.pptx", 
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
