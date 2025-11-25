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
from pptx import Presentation
from fastapi.responses import JSONResponse
from PIL import Image

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Try to import docx2pdf for Windows
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None
    logger.warning("docx2pdf not available - will use LibreOffice if available")

# Temporary upload folder
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), "file_converter_uploads")

# Lifespan context for startup/shutdown events
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

# CORS config - adjust origins in production
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

# PDF functions - lock/unlock
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
        if reader.is_encrypted:
            if not reader.decrypt(password):
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

# Office file lock/unlock (Word, Excel)
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

# PowerPoint lock/unlock
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

# ======== Photo to PDF route ========
from PIL import Image

@app.post("/photo-to-pdf")
async def photo_to_pdf(files: list[UploadFile] = File(...)):
    allowed = [".jpg", ".jpeg", ".png", ".webp"]
    if not files or len(files) == 0:
        raise HTTPException(status_code=400, detail="Please upload at least one image.")
    images = []
    try:
        for file in files:
            ext = os.path.splitext(file.filename)[1].lower()
            if ext not in allowed:
                raise HTTPException(status_code=400, detail=f"Unsupported file type: {file.filename}")
            content = await file.read()
            img = Image.open(io.BytesIO(content)).convert("RGB")
            images.append(img)
        output_pdf = io.BytesIO()
        images[0].save(
            output_pdf,
            format="PDF",
            save_all=True,
            append_images=images[1:] if len(images) > 1 else None,
        )
        output_pdf.seek(0)
        return StreamingResponse(
            output_pdf,
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=photos_converted.pdf"},
        )
    except Exception as e:
        logger.error(f"Photo to PDF Error: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to convert images: {str(e)}")

# ====================================================
#                   ROUTES: PDF → WORD
# ====================================================
@app.post("/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    """
    Convert PDF to DOCX.
    Note: Requires PyMuPDF pinned to version 1.26.4 to avoid 'Rect' attribute error.
    """
    if not validate_file_extension(file.filename, [".pdf"]):
        raise HTTPException(status_code=400, detail="Only PDF files are allowed")

    input_path = get_temp_file(file.filename)
    output_path = input_path.replace(".pdf", ".docx")

    try:
        content = await file.read()
        with open(input_path, "wb") as f:
            f.write(content)

        cv = Converter(input_path)  # uses pdf2docx which depends on PyMuPDF
        cv.convert(output_path)
        cv.close()

        with open(output_path, "rb") as f:
            output_content = f.read()

        cleanup_file(input_path)
        cleanup_file(output_path)

        return StreamingResponse(
            io.BytesIO(output_content),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f"attachment; filename={os.path.basename(file.filename).replace('.pdf', '.docx')}"
            },
        )
    except Exception as e:
        cleanup_file(input_path)
        cleanup_file(output_path)
        logger.error(f"PDF to Word conversion error: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")


# ======== Lock file route ========
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

# ======== Unlock file route ========
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
# ======== Health check routes ========
@app.get("/")
def home():
    return {
        "message": "File Converter & Locker API",
        "version": "1.0.0",
        "platform": platform.system(),
        "endpoints": {
            "pdf_to_word": "/pdf-to-word",
            "lock": "/lock",
            "unlock": "/unlock",
            "photo_to_pdf": "/photo-to-pdf"
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
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True, log_level="info")









