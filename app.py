from flask import Flask, request, jsonify
from docx import Document
import os

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_docx_to_text():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({"error": "File must be a .docx"}), 400

    # Save uploaded file temporarily
    file_path = "temp.docx"
    file.save(file_path)

    # Read DOCX content
    doc = Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])

    # Save to .txt (optional)
    with open("output.txt", "w", encoding="utf-8") as f:
        f.write(text)

    os.remove(file_path)

    return jsonify({"message": "File converted successfully", "text": text})


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Use Render’s PORT
    app.run(host="0.0.0.0", port=port)        # Bind to all interfaces
