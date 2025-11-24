# import os
# import io
# import tempfile
# import subprocess
# import platform
# import uuid
# from fastapi import FastAPI, UploadFile, File, Form
# from fastapi.responses import StreamingResponse, JSONResponse
# from fastapi.middleware.cors import CORSMiddleware

# from PyPDF2 import PdfReader, PdfWriter
# from pdf2docx import Converter
# import msoffcrypto
# import zipfile

# try:
#     from docx2pdf import convert as docx2pdf_convert
# except ImportError:
#     docx2pdf_convert = None

# # ====================================================
# #                     FASTAPI APP INIT
# # ====================================================
# app = FastAPI()

# # ====================================================
# #                 CORS CONFIGURATION
# # ====================================================
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],  # Change to your domain in production
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ====================================================
# #                 TEMP UPLOAD FOLDER
# # ====================================================
# UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), "uploads")
# os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# # ====================================================
# #              UTILITY: TEMP FILE NAME
# # ====================================================
# def get_temp_file(filename):
#     """
#     Generate a unique temporary file path to avoid overwriting files
#     """
#     uid = uuid.uuid4().hex
#     base, ext = os.path.splitext(filename)
#     return os.path.join(UPLOAD_FOLDER, f"{base}_{uid}{ext}")

# # ====================================================
# #                     PDF FUNCTIONS
# # ====================================================
# # Supported file types:
# #         - PDF      (.pdf)
# #         - Word     (.docx)
# #         - Excel    (.xlsx)
# #         - PowerPoint (.pptx)


# def lock_pdf(file_bytes, password):
#     reader = PdfReader(io.BytesIO(file_bytes))
#     writer = PdfWriter()
#     for page in reader.pages:
#         writer.add_page(page)
#     writer.encrypt(user_password=password)
#     out_stream = io.BytesIO()
#     writer.write(out_stream)
#     out_stream.seek(0)
#     return out_stream

# def unlock_pdf(file_bytes, password):
#     reader = PdfReader(io.BytesIO(file_bytes))
#     reader.decrypt(password)
#     writer = PdfWriter()
#     for page in reader.pages:
#         writer.add_page(page)
#     out_stream = io.BytesIO()
#     writer.write(out_stream)
#     out_stream.seek(0)
#     return out_stream

# # ====================================================
# #                   OFFICE FUNCTIONS
# # ====================================================
# def lock_office(file_bytes, password):
#     temp_in = io.BytesIO(file_bytes)
#     temp_out = io.BytesIO()
#     office_file = msoffcrypto.OfficeFile(temp_in)
#     office_file.load()
#     office_file.encrypt(password)
#     office_file.save(temp_out)
#     temp_out.seek(0)
#     return temp_out

# def unlock_office(file_bytes, password):
#     temp_in = io.BytesIO(file_bytes)
#     temp_out = io.BytesIO()
#     office_file = msoffcrypto.OfficeFile(temp_in)
#     office_file.load(password=password)
#     office_file.save(temp_out)
#     temp_out.seek(0)
#     return temp_out

# # ====================================================
# #                     PPT FUNCTIONS
# # ====================================================
# def lock_ppt(file_bytes, password):
#     temp_out = io.BytesIO()
#     with zipfile.ZipFile(temp_out, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
#         zf.writestr("presentation.pptx", file_bytes)
#     temp_out.seek(0)
#     return temp_out

# def unlock_ppt(file_bytes, password):
#     temp_in = io.BytesIO(file_bytes)
#     with zipfile.ZipFile(temp_in, 'r') as zf:
#         ppt_bytes = zf.read("presentation.pptx")
#     return io.BytesIO(ppt_bytes)

# # ====================================================
# #                   ROUTES: PDF → WORD
# # ====================================================
# @app.post("/pdf-to-word")
# async def pdf_to_word(file: UploadFile = File(...)):
#     """
#     Converts uploaded PDF to DOCX.
#     """
#     input_path = get_temp_file(file.filename)
#     output_path = input_path.replace(".pdf", ".docx")

#     with open(input_path, "wb") as f:
#         f.write(await file.read())

#     cv = Converter(input_path)
#     cv.convert(output_path)
#     cv.close()
#     os.remove(input_path)

#     return StreamingResponse(
#         open(output_path, "rb"),
#         media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#         headers={"Content-Disposition": f"attachment; filename={os.path.basename(output_path)}"}
#     )

# # ====================================================
# #                   ROUTES: WORD → PDF
# # ====================================================
# @app.post("/word-to-pdf")
# async def word_to_pdf(file: UploadFile = File(...)):
#     """
#     Converts uploaded DOCX to PDF.
#     """
#     input_path = get_temp_file(file.filename)
#     output_path = input_path.replace(".docx", ".pdf")

#     with open(input_path, "wb") as f:
#         f.write(await file.read())

#     try:
#         if platform.system() == "Windows" and docx2pdf_convert:
#             docx2pdf_convert(input_path, output_path)
#         else:
#             subprocess.run([
#                 "soffice", "--headless", "--convert-to", "pdf",
#                 "--outdir", UPLOAD_FOLDER, input_path
#             ], check=True)
#             output_path = os.path.join(UPLOAD_FOLDER, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")
#     except Exception as e:
#         os.remove(input_path)
#         return JSONResponse({"error": str(e)}, status_code=500)

#     os.remove(input_path)
#     return StreamingResponse(
#         open(output_path, "rb"),
#         media_type="application/pdf",
#         headers={"Content-Disposition": f"attachment; filename={os.path.basename(output_path)}"}
#     )
# # ====================================================
# #                   ROUTE: LOCK FILE
# # ====================================================
# @app.post("/lock")
# async def lock_file(file: UploadFile = File(...), password: str = Form(...)):
#     """
#     Lock PDF, DOCX, XLSX or PPTX with given password.
#     """
#     ext = os.path.splitext(file.filename)[1].lower()
#     content = await file.read()
#     try:
#         if ext == ".pdf":
#             out = lock_pdf(content, password)
#             return StreamingResponse(
#                 out,
#                 media_type="application/pdf",
#                 headers={"Content-Disposition": f"attachment; filename={file.filename.replace('.pdf','_locked.pdf')}"}
#             )
#         elif ext in [".docx", ".xlsx"]:
#             out = lock_office(content, password)
#             return StreamingResponse(
#                 out,
#                 headers={"Content-Disposition": f"attachment; filename={file.filename.replace('.docx','_locked.docx')}"}
#             )
#         elif ext == ".pptx":
#             out = lock_ppt(content, password)
#             return StreamingResponse(
#                 out,
#                 headers={"Content-Disposition": f"attachment; filename={file.filename.replace('.pptx','_locked.pptx')}"}
#             )
#         else:
#             return JSONResponse({"error": "Unsupported file type"}, status_code=400)
#     except Exception as e:
#         return JSONResponse({"error": str(e)}, status_code=500)


# # ====================================================
# #                   ROUTE: UNLOCK FILE
# # ====================================================
# @app.post("/unlock")
# async def unlock_file(file: UploadFile = File(...), password: str = Form(...)):
#     """
#     Unlock PDF, DOCX, XLSX or PPTX with given password.
#     """
#     ext = os.path.splitext(file.filename)[1].lower()
#     content = await file.read()
#     try:
#         if ext == ".pdf":
#             out = unlock_pdf(content, password)
#             return StreamingResponse(
#                 out,
#                 media_type="application/pdf",
#                 headers={"Content-Disposition": f"attachment; filename={file.filename.replace('.pdf','_unlocked.pdf')}"}
#             )
#         elif ext in [".docx", ".xlsx"]:
#             out = unlock_office(content, password)
#             return StreamingResponse(
#                 out,
#                 headers={"Content-Disposition": f"attachment; filename={file.filename.replace('_locked','_unlocked')}"}
#             )
#         elif ext == ".pptx":
#             out = unlock_ppt(content, password)
#             return StreamingResponse(
#                 out,
#                 headers={"Content-Disposition": f"attachment; filename={file.filename.replace('_locked','_unlocked')}"}
#             )
#         else:
#             return JSONResponse({"error": "Unsupported file type"}, status_code=400)
#     except Exception as e:
#         return JSONResponse({"error": str(e)}, status_code=500)

# # ====================================================
# #                   ROUTE: HOME
# # ====================================================
# @app.get("/")
# def home():
#     """
#     Simple home route for API health check.
#     """
#     return {"message": "Cross-Platform FastAPI — Lock/Unlock + PDF/Word Converter"}

# # ====================================================
# #                   APP RUNNER
# # ====================================================
# if __name__ == "__main__":
#     import uvicorn
#     uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)








import os
import io
import tempfile
import subprocess
import platform
import uuid
import logging
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
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

# ====================================================
#                  LIFESPAN CONTEXT
# ====================================================
@asynccontextmanager
async def lifespan(app: FastAPI):
    """Initialize and cleanup resources"""
    # Startup
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    logger.info(f"Upload folder created: {UPLOAD_FOLDER}")
    logger.info(f"Platform: {platform.system()}")
    yield
    # Shutdown
    logger.info("Shutting down...")

# ====================================================
#                     FASTAPI APP INIT
# ====================================================
app = FastAPI(
    title="File Converter & Locker API",
    description="Cross-platform API for PDF/Word conversion and file locking",
    version="1.0.0",
    lifespan=lifespan
)

# ====================================================
#                 CORS CONFIGURATION
# ====================================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Update with your frontend domain in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ====================================================
#                 TEMP UPLOAD FOLDER
# ====================================================
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), "file_converter_uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ====================================================
#              UTILITY FUNCTIONS
# ====================================================
def get_temp_file(filename: str) -> str:
    """Generate unique temporary file path"""
    uid = uuid.uuid4().hex
    base, ext = os.path.splitext(filename)
    return os.path.join(UPLOAD_FOLDER, f"{base}_{uid}{ext}")

def cleanup_file(filepath: str):
    """Safely remove temporary file"""
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
            logger.info(f"Cleaned up: {filepath}")
    except Exception as e:
        logger.error(f"Error cleaning up {filepath}: {e}")

def validate_file_extension(filename: str, allowed_extensions: list) -> bool:
    """Validate file extension"""
    ext = os.path.splitext(filename)[1].lower()
    return ext in allowed_extensions

# ====================================================
#                     PDF FUNCTIONS
# ====================================================
def lock_pdf(file_bytes: bytes, password: str) -> io.BytesIO:
    """Lock PDF with password"""
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        writer = PdfWriter()
        
        for page in reader.pages:
            writer.add_page(page)
        
        # Encrypt with user password and owner password
        writer.encrypt(
            user_password=password,
            owner_password=password,
            use_128bit=True
        )
        
        out_stream = io.BytesIO()
        writer.write(out_stream)
        out_stream.seek(0)
        return out_stream
    except Exception as e:
        logger.error(f"Error locking PDF: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock PDF: {str(e)}")

def unlock_pdf(file_bytes: bytes, password: str) -> io.BytesIO:
    """Unlock PDF with password"""
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

# ====================================================
#                   OFFICE FUNCTIONS (DOCX/XLSX)
# ====================================================
def lock_office_file(file_bytes: bytes, password: str) -> io.BytesIO:
    """Lock Word/Excel file with password using msoffcrypto"""
    try:
        # Read the original file
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        
        # Create OfficeFile object
        ms_file = msoffcrypto.OfficeFile(input_stream)
        
        # Encrypt with password
        ms_file.encrypt(password, output_stream)
        
        # Reset position to beginning
        output_stream.seek(0)
        return output_stream
        
    except Exception as e:
        logger.error(f"Error locking Office file: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock file: {str(e)}")

def unlock_office_file(file_bytes: bytes, password: str) -> io.BytesIO:
    """Unlock Word/Excel file with password"""
    try:
        # Read the encrypted file
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        
        # Create OfficeFile object
        ms_file = msoffcrypto.OfficeFile(input_stream)
        
        # Load password
        ms_file.load_key(password=password)
        
        # Decrypt file
        ms_file.decrypt(output_stream)
        
        # Reset position to beginning
        output_stream.seek(0)
        return output_stream
        
    except Exception as e:
        logger.error(f"Error unlocking Office file: {e}")
        raise HTTPException(status_code=401, detail="Invalid password or corrupted file")

# ====================================================
#                   POWERPOINT FUNCTIONS
# ====================================================
def lock_ppt(file_bytes: bytes, password: str) -> io.BytesIO:
    """Lock PowerPoint file using msoffcrypto"""
    try:
        # Read the original file
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        
        # Create OfficeFile object
        ms_file = msoffcrypto.OfficeFile(input_stream)
        
        # Encrypt with password
        ms_file.encrypt(password, output_stream)
        
        # Reset position to beginning
        output_stream.seek(0)
        return output_stream
        
    except Exception as e:
        logger.error(f"Error locking PowerPoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to lock PowerPoint: {str(e)}")

def unlock_ppt(file_bytes: bytes, password: str) -> io.BytesIO:
    """Unlock PowerPoint file"""
    try:
        # Read the encrypted file
        input_stream = io.BytesIO(file_bytes)
        output_stream = io.BytesIO()
        
        # Create OfficeFile object
        ms_file = msoffcrypto.OfficeFile(input_stream)
        
        # Load password
        ms_file.load_key(password=password)
        
        # Decrypt file
        ms_file.decrypt(output_stream)
        
        # Reset position to beginning
        output_stream.seek(0)
        return output_stream
        
    except Exception as e:
        logger.error(f"Error unlocking PowerPoint: {e}")
        raise HTTPException(status_code=401, detail="Invalid password or corrupted file")

# ====================================================
#                   ROUTES: PDF → WORD
# ====================================================
@app.post("/pdf-to-word")
async def pdf_to_word(file: UploadFile = File(...)):
    """Convert PDF to DOCX"""
    if not validate_file_extension(file.filename, [".pdf"]):
        raise HTTPException(status_code=400, detail="Only PDF files are allowed")
    
    input_path = get_temp_file(file.filename)
    output_path = input_path.replace(".pdf", ".docx")
    
    try:
        # Save uploaded file
        content = await file.read()
        with open(input_path, "wb") as f:
            f.write(content)
        
        # Convert PDF to DOCX
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()
        
        # Read output file
        with open(output_path, "rb") as f:
            output_content = f.read()
        
        # Cleanup
        cleanup_file(input_path)
        cleanup_file(output_path)
        
        return StreamingResponse(
            io.BytesIO(output_content),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f"attachment; filename={os.path.basename(file.filename).replace('.pdf', '.docx')}"
            }
        )
    except Exception as e:
        cleanup_file(input_path)
        cleanup_file(output_path)
        logger.error(f"PDF to Word conversion error: {e}")
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

# ====================================================
#                   ROUTES: WORD → PDF
# ====================================================
@app.post("/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    """Convert DOCX to PDF"""
    if not validate_file_extension(file.filename, [".docx", ".doc"]):
        raise HTTPException(status_code=400, detail="Only Word files are allowed")
    
    input_path = get_temp_file(file.filename)
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(UPLOAD_FOLDER, f"{base_name}.pdf")
    
    try:
        # Save uploaded file
        content = await file.read()
        with open(input_path, "wb") as f:
            f.write(content)
        
        # Convert based on platform
        if platform.system() == "Windows" and docx2pdf_convert:
            # Use docx2pdf on Windows
            docx2pdf_convert(input_path, output_path)
        else:
            # Use LibreOffice on Linux/Mac
            result = subprocess.run([
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", UPLOAD_FOLDER, input_path
            ], capture_output=True, text=True, timeout=60)
            
            if result.returncode != 0:
                raise Exception(f"LibreOffice conversion failed: {result.stderr}")
        
        # Verify output file exists
        if not os.path.exists(output_path):
            raise Exception("Output file was not created")
        
        # Read output file
        with open(output_path, "rb") as f:
            output_content = f.read()
        
        # Cleanup
        cleanup_file(input_path)
        cleanup_file(output_path)
        
        return StreamingResponse(
            io.BytesIO(output_content),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename={os.path.basename(file.filename).replace('.docx', '.pdf').replace('.doc', '.pdf')}"
            }
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

# ====================================================
#                   ROUTE: LOCK FILE
# ====================================================
@app.post("/lock")
async def lock_file(file: UploadFile = File(...), password: str = Form(...)):
    """Lock PDF, DOCX, XLSX, or PPTX with password"""
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

# ====================================================
#                   ROUTE: UNLOCK FILE
# ====================================================
@app.post("/unlock")
async def unlock_file(file: UploadFile = File(...), password: str = Form(...)):
    """Unlock PDF, DOCX, XLSX, or PPTX with password"""
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

# ====================================================
#                   ROUTE: HEALTH CHECK
# ====================================================
@app.get("/")
def home():
    """API health check"""
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
    """Detailed health check"""
    return {
        "status": "healthy",
        "platform": platform.system(),
        "docx2pdf_available": docx2pdf_convert is not None,
        "upload_folder": UPLOAD_FOLDER
    }

# ====================================================
#                   APP RUNNER
# ====================================================
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )