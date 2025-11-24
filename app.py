import os
import io
import tempfile
import subprocess
import platform
import uuid
import logging
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter
import msoffcrypto
from openpyxl import load_workbook
from pptx import Presentation

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Try to import docx2pdf for Windows
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None
    logger.warning("docx2pdf not available - will use LibreOffice if available")

# Lifespan context for startup/shutdown
@asynccontextmanager
async def lifespan(app: FastAPI):
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    logger.info(f"Upload folder created: {UPLOAD_FOLDER}")
    logger.info(f"Platform: {platform.system()}")
    yield
    logger.info("Shutting down...")

# Initialize FastAPI app
app = FastAPI(
    title="File Converter & Locker API",
    description="Cross-platform API for PDF/Word conversion and file locking",
    version="1.0.0",
    lifespan=lifespan
)

# Enable CORS for all origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Temporary upload folder
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), "file_converter_uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Utility functions
def get_temp_file(filename: str) -> str:
    uid = uuid.uuid4().hex
    base, ext = os.path.splitext(filename)
    return os.path.join(UPLOAD_FOLDER, f"{base}_{uid}{ext}")

def cleanup_file(filepath: str):
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
            logger.info(f"Cleaned up: {filepath}")
    except Exception as e:
        logger.error(f"Error cleaning up {filepath}: {e}")

def validate_file_extension(filename: str, allowed_extensions: list) -> bool:
    ext = os.path.splitext(filename)[1].lower()
    return ext in allowed_extensions

# PDF Lock and Unlock functions (same as before)
def lock_pdf(file_bytes: bytes, password: str) -> io.BytesIO:
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(user_password=password, owner_password=password, use_128bit=True)
        out_stream = io.BytesIO()
        writer.write(out_stream)
        out_stream.seek(0)
        return out_stream
    except Exception as e:
        logger.error(f"Error locking PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock PDF: {str(e)}")

def unlock_pdf(file_bytes: bytes, password: str) -> io.BytesIO:
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        if reader.is_encrypted and not reader.decrypt(password):
            raise HTTPException(status_code=401, detail="Invalid password")
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        out_stream = io.BytesIO()
        writer.write(out_stream)
        out_stream.seek(0)
        return out_stream
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error unlocking PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to unlock PDF: {str(e)}")

# Office file lock/unlock using msoffcrypto (same as before)
def lock_office_file(file_bytes: bytes, password: str) -> io.BytesIO:
    try:
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        ms_file = msoffcrypto.OfficeFile(input_stream)
        ms_file.encrypt(password, output_stream)
        output_stream.seek(0)
        return output_stream
    except Exception as e:
        logger.error(f"Error locking Office file: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock file: {str(e)}")

def unlock_office_file(file_bytes: bytes, password: str) -> io.BytesIO:
    try:
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        ms_file = msoffcrypto.OfficeFile(input_stream)
        ms_file.load_key(password=password)
        ms_file.decrypt(output_stream)
        output_stream.seek(0)
        return output_stream
    except Exception as e:
        logger.error(f"Error unlocking Office file: {e}")
        raise HTTPException(status_code=401, detail="Invalid password or corrupted file")

# PowerPoint lock/unlock (same as before)
def lock_ppt(file_bytes: bytes, password: str) -> io.BytesIO:
    try:
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        ms_file = msoffcrypto.OfficeFile(input_stream)
        ms_file.encrypt(password, output_stream)
        output_stream.seek(0)
        return output_stream
    except Exception as e:
        logger.error(f"Error locking PowerPoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock PowerPoint: {str(e)}")

def unlock_ppt(file_bytes: bytes, password: str) -> io.BytesIO:
    try:
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        ms_file = msoffcrypto.OfficeFile(input_stream)
        ms_file.load_key(password=password)
        ms_file.decrypt(output_stream)
        output_stream.seek(0)
        return output_stream
    except Exception as e:
        logger.error(f"Error unlocking PowerPoint: {e}")
        raise HTTPException(status_code=401, detail="Invalid password or corrupted file")

# Route: PDF to Word (same as before)
@app.post("/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    if not validate_file_extension(file.filename, [".pdf"]):
        raise HTTPException(status_code=400, detail="Only PDF files are allowed")
    input_path = get_temp_file(file.filename)
    output_path = input_path.replace(".pdf", ".docx")
    try:
        content = await file.read()
        with open(input_path, "wb") as f:
            f.write(content)
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()
        with open(output_path, "rb") as f:
            output_content = f.read()
        cleanup_file(input_path)
        cleanup_file(output_path)
        return StreamingResponse(
            io.BytesIO(output_content),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={os.path.basename(file.filename).replace('.pdf', '.docx')}"}
        )
    except Exception as e:
        cleanup_file(input_path)
        cleanup_file(output_path)
        logger.error(f"PDF to Word conversion error: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

# Replaced Route: Word to PDF
@app.post("/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    if not validate_file_extension(file.filename, [".docx", ".doc"]):
        raise HTTPException(status_code=400, detail="Only Word files are allowed")
    input_path = get_temp_file(file.filename)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(UPLOAD_FOLDER, f"{base_name}.pdf")
    content = await file.read()
    with open(input_path, "wb") as f:
        f.write(content)
    try:
        command = [
            "libreoffice", "--headless", "--convert-to", "pdf",
            input_path, "--outdir", UPLOAD_FOLDER
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            raise Exception(f"LibreOffice conversion failed: {result.stderr}")
        if not os.path.exists(output_path):
            raise Exception("Output file was not created")
        with open(output_path, "rb") as f:
            output_content = f.read()
        return StreamingResponse(
            io.BytesIO(output_content),
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={base_name}.pdf"}
        )
    except subprocess.TimeoutExpired:
        cleanup_file(input_path)
        cleanup_file(output_path)
        raise HTTPException(status_code=500, detail="Conversion timeout - file too large")
    except Exception as e:
        cleanup_file(input_path)
        cleanup_file(output_path)
        logger.error(f"Word to PDF conversion error: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

# Route: Lock file
@app.post("/lock")
async def lock_file(file: UploadFile = File(...), password: str = Form(...)):
    if not password or len(password) < 4:
        raise HTTPException(status_code=400, detail="Password must be at least 4 characters")
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".pdf", ".docx", ".xlsx", ".pptx"]:
        raise HTTPException(status_code=400, detail="Unsupported file type. Allowed: PDF, DOCX, XLSX, PPTX")
    try:
        content = await file.read()
        if ext == ".pdf":
            out = lock_pdf(content, password)
            media_type = "application/pdf"
            new_filename = file.filename.replace(".pdf", "_locked.pdf")
        elif ext in [".docx", ".xlsx"]:
            out = lock_office_file(content, password)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == ".docx" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            new_filename = file.filename.replace(ext, f"_locked{ext}")
        elif ext == ".pptx":
            out = lock_ppt(content, password)
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            new_filename = file.filename.replace(".pptx", "_locked.pptx")
        return StreamingResponse(
            out,
            media_type=media_type,
            headers={"Content-Disposition": f"attachment; filename={new_filename}"}
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Lock file error: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock file: {str(e)}")

# Route: Unlock file
@app.post("/unlock")
async def unlock_file(file: UploadFile = File(...), password: str = Form(...)):
    if not password:
        raise HTTPException(status_code=400, detail="Password is required")
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in [".pdf", ".docx", ".xlsx", ".pptx"]:
        raise HTTPException(status_code=400, detail="Unsupported file type")
    try:
        content = await file.read()
        if ext == ".pdf":
            out = unlock_pdf(content, password)
            media_type = "application/pdf"
            new_filename = file.filename.replace("_locked", "_unlocked")
        elif ext in [".docx", ".xlsx"]:
            out = unlock_office_file(content, password)
            media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if ext == ".docx" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            new_filename = file.filename.replace("_locked", "_unlocked")
        elif ext == ".pptx":
            out = unlock_ppt(content, password)
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            new_filename = file.filename.replace("_locked", "_unlocked")
        return StreamingResponse(
            out,
            media_type=media_type,
            headers={"Content-Disposition": f"attachment; filename={new_filename}"}
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unlock file error: {e}")
        raise HTTPException(status_code=500, detail="Invalid password or corrupted file")

# Health check route
@app.get("/")
def home():
    return {
        "message": "File Converter & Locker API",
        "version": "1.0.0",
        "platform": platform.system(),
        "endpoints": {
            "pdf_to_word": "/pdf-to-word",
            "word_to_pdf": "/word-to-pdf",
            "lock": "/lock",
            "unlock": "/unlock"
        }
    }

@app.get("/health")
def health_check():
    return {
        "status": "healthy",
        "platform": platform.system(),
        "docx2pdf_available": docx2pdf_convert is not None,
        "upload_folder": UPLOAD_FOLDER
    }

# Run app
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )
