import os
import uuid
import base64

from flask import Flask, request, render_template, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pptx import Presentation

# Enums for shapes/fills
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL

app = Flask(__name__)
app.secret_key = "YOUR_SECRET_KEY_HERE"  # Replace with something secure

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"pptx"}

def allowed_file(filename):
    """Check if the uploaded file has a .pptx extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET"])
def index():
    """Show an upload form."""
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_pptx():
    """Handle file upload, parse PPTX, and render results."""
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

def parse_pptx(filepath):
    """
    Parse slides in a comprehensive way:
      - Master & layout images
      - Slide backgrounds
      - Group shapes, picture fills
      - Nested bullet lists
      - Hyperlinks (YouTube)
    """
    prs = Presentation(filepath)
    slides_data = []

    # 1) Gather images from the SLIDE MASTER
    master_images = []
    if prs.slide_master:
        # shapes on master
        for shape in prs.slide_master.shapes:
            handle_shape_for_images(shape, master_images)
        # background on master
        handle_background_for_images(prs.slide_master, master_images)

    # 2) Build a dict mapping LAYOUT INDEX -> images
    layout_images_map = {}
    for i, layout in enumerate(prs.slide_layouts):
        layout_img_list = []
        # shapes in layout
        for shape in layout.shapes:
            handle_shape_for_images(shape, layout_img_list)
        # background in layout
        handle_background_for_images(layout, layout_img_list)
        layout_images_map[i] = layout_img_list

    # 3) Parse each SLIDE
    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        title = None
        images = []
        youtube_links = []
        text_html = ""

        # Merge in master images
        images.extend(master_images)

        # Merge in layout images for this slide
        # (Find the layout's index in prs.slide_layouts)
        try:
            layout_index = prs.slide_layouts.index(slide.slide_layout)
            images.extend(layout_images_map[layout_index])
        except ValueError:
            # If slide_layout isn't in prs.slide_layouts for some reason, skip
            pass

        # Also parse the slide's own background
        handle_background_for_images(slide, images)

        # Now parse shapes on the slide itself
        for shape in slide.shapes:
            # images from shapes (group/picture fill)
            handle_shape_for_images(shape, images)

            # If shape is a TITLE placeholder
            if shape.is_placeholder and shape.placeholder_format.type == 1:
                title = shape.text

            # If shape has text, parse bullet paragraphs & detect YouTube links
            if shape.has_text_frame:
                # nested bullet lists
                text_html += paragraphs_to_nested_lists(shape.text_frame.paragraphs)
                # check hyperlinks
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.hyperlink and run.hyperlink.address:
                            link = run.hyperlink.address
                            if "youtube.com" in link or "youtu.be" in link:
                                youtube_links.append(link)

        # If no title, default
        if not title:
            title = f"Slide {slide_num}"

        # Deduplicate images (preserve order)
        images = list(dict.fromkeys(images))

        slides_data.append({
            "title": title,
            "slide_number": slide_num,
            "text_html": text_html,
            "images": images,
            "youtube_links": youtube_links
        })

    return slides_data

# --------------------------------------------
# HELPER: Handle shape images (group shapes, fills)
# --------------------------------------------
def handle_shape_for_images(shape, images_list):
    """
    Checks if a shape has an image or is a group shape, or a fill picture.
    Appends base64 data URIs to 'images_list'.
    """
    # If shape is a PICTURE or LINKED_PICTURE
    if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE):
        if getattr(shape, "image", None):
            embed_image_as_base64(shape.image, images_list)

    # If shape is a GROUP, recurse into subshapes
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for subshape in shape.shapes:
            handle_shape_for_images(subshape, images_list)

    # If shape has a picture fill
    if getattr(shape, "fill", None) and shape.fill.type == MSO_FILL.PICTURE:
        if getattr(shape.fill.picture, "blob", None):
            embed_image_as_base64(shape.fill.picture, images_list)

def handle_background_for_images(slide_master, master_images):
    """
    Extracts background images from slide master and adds them to the images list.
    """
    for layout in slide_master.slide_layouts:
        for shape in layout.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:  # Check for grouped shapes
                for subshape in shape.shapes:
                    if hasattr(subshape, "image") and subshape.image:
                        base64data = base64.b64encode(subshape.image.blob).decode("utf-8")
                        mime_type = "image/" + subshape.image.ext.lower()
                        data_uri = f"data:{mime_type};base64,{base64data}"
                        master_images.append(data_uri)

            # Check for a picture in the shape
            if hasattr(shape, "image") and shape.image:
                base64data = base64.b64encode(shape.image.blob).decode("utf-8")
                mime_type = "image/" + shape.image.ext.lower()
                data_uri = f"data:{mime_type};base64,{base64data}"
                master_images.append(data_uri)

            # Check for background fill format with a picture
            if hasattr(shape.fill, "picture") and shape.fill.picture:
                picture = shape.fill.picture
                if hasattr(picture, "blob"):
                    base64data = base64.b64encode(picture.blob).decode("utf-8")
                    mime_type = "image/" + picture.ext.lower()
                    data_uri = f"data:{mime_type};base64,{base64data}"
                    master_images.append(data_uri)


    ext = image_obj.ext.lower()
    if ext in ["jpg", "jpeg"]:
        mime_type = "image/jpeg"
    elif ext == "png":
        mime_type = "image/png"
    else:
        mime_type = f"image/{ext}"

    data_uri = f"data:{mime_type};base64,{base64data}"
    images_list.append(data_uri)

# --------------------------------------------
# BULLET LOGIC: Nested <ul>
# --------------------------------------------
def paragraphs_to_nested_lists(paragraphs):
    """
    Convert paragraphs with paragraph.level=0,1,2,... to nested <ul> lists.
    """
    text_html = ""
    level_stack = []

    for paragraph in paragraphs:
        runs_html = build_runs_html(paragraph.runs)
        level = paragraph.level  # 0,1,2... or -1 if none

        if not runs_html.strip():
            continue

        # If new bullet level is deeper
        while len(level_stack) < (level + 1):
            text_html += "<ul>"
            level_stack.append(True)

        # If new bullet level is shallower
        while len(level_stack) > (level + 1):
            text_html += "</ul>"
            level_stack.pop()

        text_html += f"<li>{runs_html}</li>"

    # Close any leftover <ul>
    while level_stack:
        text_html += "</ul>"
        level_stack.pop()

    return text_html

def build_runs_html(runs):
    """Apply bold/italic to each run."""
    runs_html = ""
    for run in runs:
        run_text = run.text.replace("<", "&lt;").replace(">", "&gt;")
        if getattr(run, "bold", False):
            run_text = f"<strong>{run_text}</strong>"
        if getattr(run, "italic", False):
            run_text = f"<em>{run_text}</em>"
        runs_html += run_text
    return runs_html

# --------------------------------------------
# FLASK MAIN
# --------------------------------------------
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5001)

