# app.py
import os
import tempfile
from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename
from pptx import Presentation
from io import BytesIO

# text extraction
from pdfminer.high_level import extract_text as extract_pdf_text
from docx import Document

app = Flask(__name__)

# ---------- HELPERS ----------
def save_temp_file(uploaded_file):
    filename = secure_filename(uploaded_file.filename)
    fd, path = tempfile.mkstemp(suffix=os.path.splitext(filename)[1])
    with os.fdopen(fd, 'wb') as tmp:
        tmp.write(uploaded_file.read())
    return path, filename

# ---------- ENDPOINT: extract_text ----------
@app.route('/extract_text', methods=['POST'])
def extract_text():
    if 'file' not in request.files:
        return jsonify({"error":"No file part"}), 400
    f = request.files['file']
    path, filename = save_temp_file(f)
    text = ""
    try:
        if filename.lower().endswith('.pdf'):
            text = extract_pdf_text(path)
        elif filename.lower().endswith('.docx'):
            doc = Document(path)
            text = "\n".join([p.text for p in doc.paragraphs])
        else:
            return jsonify({"error":"Unsupported file type"}), 400
    except Exception as e:
        return jsonify({"error":"Extraction failed","detail": str(e)}), 500
    finally:
        try:
            os.remove(path)
        except:
            pass
    return jsonify({"text": text})

# ---------- ENDPOINT: generate ----------
@app.route('/generate', methods=['POST'])
def generate():
    # Expect JSON with 'template' and 'slides'
    data = request.get_json(force=True, silent=True)
    if not data:
        return jsonify({"error":"Invalid JSON body"}), 400
    template_name = data.get('template')
    slides = data.get('slides', [])
    if not template_name or not isinstance(slides, list):
        return jsonify({"error":"template and slides required"}), 400

    # Template should be present in folder ./templates/
    template_path = os.path.join(os.getcwd(), 'templates', template_name)
    if not os.path.exists(template_path):
        return jsonify({"error": f"Plantilla {template_name} no encontrada."}), 404

    try:
        prs = Presentation(template_path)
        # Strategy: use layout 1 for content slides, but adapt if not available
        layout = None
        if len(prs.slide_layouts) > 1:
            layout = prs.slide_layouts[1]
        else:
            layout = prs.slide_layouts[0]

        # Append slides according to slides list
        for s in slides:
            title = s.get('title', '')
            body = s.get('body', '')
            slide = prs.slides.add_slide(layout)
            # try fill title
            try:
                if slide.shapes.title:
                    slide.shapes.title.text = title
            except Exception:
                pass
            # try fill first content placeholder or first text_frame that's not title
            filled = False
            for shape in slide.shapes:
                if shape.has_text_frame:
                    # skip if it is title shape
                    try:
                        if shape is slide.shapes.title:
                            continue
                    except Exception:
                        pass
                    shape.text_frame.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = body
                    filled = True
                    break
            # if nothing filled, create a textbox
            if not filled:
                from pptx.util import Inches, Pt
                left = Inches(1)
                top = Inches(1.5)
                width = Inches(8)
                height = Inches(4.5)
                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                tf.text = body

        out = BytesIO()
        prs.save(out)
        out.seek(0)
        return send_file(out,
                         mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                         as_attachment=True,
                         download_name='presentation_generated.pptx')
    except Exception as e:
        return jsonify({"error":"Generation failed","detail": str(e)}), 500

# ---------- Root for health check ----------
@app.route('/')
def home():
    return jsonify({"message":"PPT Generator Flask App is running!"})

if __name__ == '__main__':
    # Render requires using the PORT env var
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
