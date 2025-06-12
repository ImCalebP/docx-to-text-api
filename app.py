from flask import Flask, request, jsonify
from docx import Document
from docx_properties import core_properties
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
        # Extract text
        doc = Document(docx_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])

        # Metadata
        props = core_properties(docx_path)
        metadata = {
            "numpages": 1,  # Estimating 1 (docx doesn't track this)
            "numrender": 1,
            "info": {
                "PDFFormatVersion": None,
                "Language": None,
                "EncryptFilterName": None,
                "IsLinearized": False,
                "IsAcroFormPresent": False,
                "IsXFAPresent": False,
                "IsCollectionPresent": False,
                "IsSignaturesPresent": False,
                "CreationDate": props.created.isoformat() if props.created else None,
                "Creator": props.creator,
                "ModDate": props.modified.isoformat() if props.modified else None,
                "Custom": {
                    "Application": props.last_modified_by
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
