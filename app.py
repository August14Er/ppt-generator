from flask import Flask, request, send_file, jsonify
from pptx import Presentation
import os
from io import BytesIO

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate_ppt():
    data = request.get_json()
    template_name = data.get("template", "default.pptx")
    slides = data.get("slides", [])

    template_path = os.path.join("templates", template_name)
    if not os.path.exists(template_path):
        return jsonify({"error": f"Plantilla {template_name} no encontrada."}), 404

    prs = Presentation(template_path)
    for slide_data in slides:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        body = slide.placeholders[1]
        title.text = slide_data.get("title", "")
        body.text = slide_data.get("body", "")

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="generated.pptx")

@app.route('/')
def home():
    return "Flask PPT Generator funcionando."

if __name__ == '__main__':
    # Render asigna automáticamente un puerto en la variable PORT
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
