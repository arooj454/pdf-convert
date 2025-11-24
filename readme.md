curl -X POST "http://127.0.0.1:8000/word-to-pdf" \
  -F "file=@D:/Arooj_work/TS_Technology_Office_work/TS_Technology_Office_work/PDF compute/ok_converted.docx" \
  --output "ok_converted.pdf"


curl -X POST "http://127.0.0.1:8000/pdf-to-word" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@C:/Users/hp/Downloads/ok.pdf" \
  --output "ok_converted.docx"




curl -X POST "http://127.0.0.1:8000/lock" \
  -F "file=@C:/Users/hp/Downloads/ok.pdf" \
  -F "password=1234" \
  --output "C:/Users/hp/Downloads/ok_locked.pdf"




curl -X POST "http://127.0.0.1:8000/unlock" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@D:/Arooj_work/TS_Technology_Office_work/TS_Technology_Office_work/PDF compute/ok_locked.pdf" \
  -F "password=1234" \
  --output "ok_unlocked.pdf"




fastapi
uvicorn
PyPDF2
pdf2docx
python-docx
msoffcrypto-tool
openpyxl
python-pptx
pypandoc
PyYAML
# pdf-convert
