# app.py
from flask import Flask, request, send_file
#Import "flask" could not be resolvedPylancereportMissingImports)
from pptx import Presentation
#Import "pptx" could not be resolvedPylancereportMissingImports)
import tempfile, os, json

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate_ppt():
    data = request.json

    # Datos del JSON
    content = data.get('slides', [])
    template_name = data.get('template', 'default.pptx')

    # Ruta a plantilla (en la misma carpeta)
    template_path = os.path.join(os.getcwd(), 'templates', template_name)

    if not os.path.exists(template_path):
        return {"error": f"Plantilla {template_name} no encontrada."}, 404

    prs = Presentation(template_path)

    # Rellenar las diapositivas con los datos del documento
    for slide_data in content:
        title = slide_data.get('title', '')
        body = slide_data.get('body', '')

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = title

        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = body

    # Guardar temporalmente
    output_path = tempfile.mktemp(suffix=".pptx")
    prs.save(output_path)

    return send_file(output_path, as_attachment=True, download_name="presentation.pptx")

@app.route('/')
def home():
    return {"message": "PPT Generator Flask App is running!"}

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
