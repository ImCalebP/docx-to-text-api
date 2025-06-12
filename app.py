from flask import Flask, request, jsonify
from docx import Document
import os
import uuid

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_docx_to_text():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({"error": "File must be a .docx"}), 400

    # Save file with unique name
    unique_id = str(uuid.uuid4())
    docx_path = f"{unique_id}.docx"
    file.save(docx_path)

    try:
        doc = Document(docx_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])
        props = doc.core_properties

        metadata = {
            "numpages": 1,
            "numrender": 1,
            "info": {
                "PDFFormatVersion": None,
                "Language": props.language if props.language else None,
                "EncryptFilterName": None,
                "IsLinearized": False,
                "IsAcroFormPresent": False,
                "IsXFAPresent": False,
                "IsCollectionPresent": False,
                "IsSignaturesPresent": False,
                "CreationDate": props.created.isoformat() if props.created else None,
                "Creator": props.author,
                "ModDate": props.modified.isoformat() if props.modified else None,
                "Custom": {
                    "Title": props.title,
                    "Subject": props.subject,
                    "Category": props.category,
                    "Comments": props.comments,
                },
                "Producer": "python-docx",
                "Trapped": {
                    "name": "False"
                }
            },
            "text": text,
            "version": "1.0.0"
        }

        return jsonify(metadata)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(docx_path):
            os.remove(docx_path)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
