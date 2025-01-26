import os
import uuid
import base64
from flask import Flask, request, render_template, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.exc import PackageNotFoundError
from zipfile import BadZipFile

os.makedirs("static/slide_images", exist_ok=True)

app = Flask(__name__)
app.secret_key = "YOUR_SECRET_KEY_HERE"  # Replace with a secure key

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"pptx"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_pptx():
    if "pptx_file" not in request.files:
        flash("No file part in the request.")
        return redirect(url_for("index"))

    file = request.files["pptx_file"]
    if file.filename == "":
        flash("No selected file.")
        return redirect(url_for("index"))

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)

        slides_data = parse_pptx(filepath)
        return render_template("results.html", slides_data=slides_data)
    else:
        flash("Invalid file type. Please upload a .pptx file.")
        return redirect(url_for("index"))

def embed_image_as_base64(image_obj, images_list):
    """Convert raw image blob to base64 data URI, append to images_list."""
    blob = getattr(image_obj, "blob", None)
    if not blob:
        return
    base64data = base64.b64encode(blob).decode("utf-8")

    ext = image_obj.ext.lower()
    if ext in ["jpg", "jpeg"]:
        mime_type = "image/jpeg"
    elif ext == "png":
        mime_type = "image/png"
    else:
        mime_type = f"image/{ext}"

    data_uri = f"data:{mime_type};base64,{base64data}"
    images_list.append(data_uri)
bullet_styles = {
    0: "•",   # Level 1
    1: "○",   # Level 2
    2: "▪",   # Level 3
    3: "▫",   # Level 4
    4: "‣",   # Level 5
    5: "⁃",   # Level 6
    6: "✦",   # Level 7
    7: "➢",   # Level 8
}

def parse_pptx(filepath):
    try:
        prs = Presentation(filepath)
    except (PackageNotFoundError, BadZipFile):
        flash("Uploaded file is not a valid PowerPoint or is corrupted.")
        return redirect(url_for("index"))

    slides_data = []

    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        title = None
        youtube_links = []
        images = []
        text_html = ""

        for shape in slide.shapes:
            # Handle title
            if shape.is_placeholder and shape.placeholder_format.type == 1:  # TITLE
                title = shape.text

            # Handle text
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    bullet_level = paragraph.level
                    runs_html = ""

                    for run in paragraph.runs:
                        run_text = run.text.replace("<", "&lt;").replace(">", "&gt;")
                        if hasattr(run, "bold") and run.bold:
                            run_text = f"<strong>{run_text}</strong>"
                        if hasattr(run, "italic") and run.italic:
                            run_text = f"<em>{run_text}</em>"
                        runs_html += run_text

                    if runs_html.strip():
                        bullet_symbol = bullet_styles.get(bullet_level, "•")  # Default to "•" if level not in bullet_styles
                        text_html += f"<li style='margin-left:{20 * bullet_level}px'>{bullet_symbol} {runs_html}</li>"


            # Handle images
            if hasattr(shape, "image") and shape.image:
                embed_image_as_base64(shape.image, images)

            # Handle grouped shapes
            if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
                for subshape in shape.shapes:
                    if hasattr(subshape, "image") and subshape.image:
                        embed_image_as_base64(subshape.image, images)

            # Handle background images
            if hasattr(shape, "fill") and hasattr(shape.fill, "picture") and shape.fill.picture:
                embed_image_as_base64(shape.fill.picture, images)

        if not title:
            title = f"Slide {slide_num}"

        slides_data.append({
            "title": title,
            "slide_number": slide_num,
            "text_html": f"<ul>{text_html}</ul>" if text_html else "",
            "images": images,
            "youtube_links": youtube_links
        })

    return slides_data

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5001)

