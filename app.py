from flask import Flask, request, send_file
from flask_cors import CORS
import subprocess
import os
import tempfile

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

def convert_to_pdf(input_path):
    output_path = input_path.rsplit('.', 1)[0] + ".pdf"
    command = ['libreoffice', '--headless', '--convert-to', 'pdf', input_path, '--outdir', os.path.dirname(input_path)]
    subprocess.run(command, check=True)
    return output_path

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('file')
    if not file:
        return "No file uploaded", 400

    with tempfile.TemporaryDirectory() as tmpdirname:
        input_path = os.path.join(tmpdirname, file.filename)
        file.save(input_path)
        try:
            pdf_path = convert_to_pdf(input_path)
            return send_file(pdf_path, as_attachment=True)
        except subprocess.CalledProcessError:
            return "Conversion failed", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
