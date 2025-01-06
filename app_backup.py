import os
import uuid
import base64

from flask import Flask, request, render_template, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pptx import Presentation

# Enums for shapes, fills
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL

app = Flask(__name__)
app.secret_key = "YOUR_SECRET_KEY_HERE"  # replace with something secure

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {"pptx"}

def allowed_file(filename):
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
    Parse the PPTX in a robust way:
      - Handle slides, slide backgrounds, slide masters, layouts
      - Extract images (group shapes, picture fills, background images)
      - Build multi-level bullet lists
      - Gather hyperlink (YouTube) links
    """
    prs = Presentation(filepath)
    slides_data = []

    # Step 1: Pre-scan master & layout shapes in case slides use them
    # We'll create a dict mapping slide layout to a list of base64 images,
    # plus the master to a list of images, so we can merge them into each slide if used.
    master_images = []
    layout_images_map = {}

    # Parse the master slide for images
    if prs.slide_master:
        for shape in prs.slide_master.shapes:
            # recursively gather images
            handle_shape_for_images(shape, master_images)
        # Also check master background
        handle_background_for_images(prs.slide_master, master_images)

    # Parse each layout
    for layout in prs.slide_layouts:
        layout_img_list = []
        # shapes
        for shape in layout.shapes:
            handle_shape_for_images(shape, layout_img_list)
        # background
        handle_background_for_images(layout, layout_img_list)

        layout_images_map[layout] = layout_img_list

    # Step 2: Parse each actual slide
    for i, slide in enumerate(prs.slides):
        slide_num = i + 1
        title = None
        images = []
        youtube_links = []
        text_html = ""

        # A) Merge in images from the master
        images.extend(master_images)

        # B) Merge in images from the layout used by this slide (if any)
        if slide.slide_layout in layout_images_map:
            images.extend(layout_images_map[slide.slide_layout])

        # C) Check the slide's own background
        handle_background_for_images(slide, images)

        # D) Now parse all shapes on this slide
        for shape in slide.shapes:
            # 1) Extract images from shapes (including group/picture fills)
            handle_shape_for_images(shape, images)

            # 2) If shape is a TITLE placeholder
            if shape.is_placeholder and shape.placeholder_format.type == 1:
                title = shape.text

            # 3) If shape has text, parse bullet paragraphs + hyperlinks
            if shape.has_text_frame:
                # bullets
                text_html += paragraphs_to_nested_lists(shape.text_frame.paragraphs)

                # hyperlinks
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.hyperlink and run.hyperlink.address:
                            link = run.hyperlink.address
                            if "youtube.com" in link or "youtu.be" in link:
                                youtube_links.append(link)

        # E) If no title, default
        if not title:
            title = f"Slide {slide_num}"

        # F) Deduplicate images (some might appear in master + slide)
        images = list(dict.fromkeys(images))  # preserve order, remove duplicates

        slides_data.append({
            "title": title,
            "slide_number": slide_num,
            "text_html": text_html,
            "images": images,
            "youtube_links": youtube_links
        })

    return slides_data

# ----------------------------------------------------------------
# IMAGE HANDLING
# ----------------------------------------------------------------
def handle_shape_for_images(shape, images_list):
    """
    Recursively checks if a shape (or subshapes) has an image:
      - PICTURE, LINKED_PICTURE
      - GROUP shapes
      - Picture fills (AUTO_SHAPE)
    Appends base64 data URIs to images_list.
    """
    # If shape is a picture/linked picture
    if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE):
        if getattr(shape, "image", None):
            embed_image_as_base64(shape.image, images_list)

    # If shape is a group, go into subshapes
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for subshape in shape.shapes:
            handle_shape_for_images(subshape, images_list)

    # If shape has a picture fill
    if getattr(shape, "fill", None) and shape.fill.type == MSO_FILL.PICTURE:
        if getattr(shape.fill.picture, "blob", None):
            embed_image_as_base64(shape.fill.picture, images_list)

def handle_background_for_images(slide_or_layout, images_list):
    """
    Some slides/layouts/masters have a 'background' property with a fill.
    If it's a picture fill, embed that too.
    """
    # Not all objects have .background, so guard
    bg = getattr(slide_or_layout, "background", None)
    if not bg or not bg.fill:
        return

    if bg.fill.type == MSO_FILL.PICTURE:
        if getattr(bg.fill.picture, "blob", None):
            embed_image_as_base64(bg.fill.picture, images_list)

def embed_image_as_base64(image_obj, images_list):
    """
    Convert raw blob to base64 data URI and append to images_list.
    Avoid duplicates if possible (you can handle that at the end).
    """
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

# ----------------------------------------------------------------
# BULLETS: NESTED <ul> LISTS
# ----------------------------------------------------------------
def paragraphs_to_nested_lists(paragraphs):
    """
    Convert paragraphs with paragraph.level=0,1,2,... into nested <ul> lists.
    """
    text_html = ""
    level_stack = []  # We'll push/pop <ul> for each bullet level

    for paragraph in paragraphs:
        runs_html = build_runs_html(paragraph.runs)
        level = paragraph.level  # 0,1,2,... or -1

        # Skip empty paragraphs
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

    # Close any remaining <ul>
    while level_stack:
        text_html += "</ul>"
        level_stack.pop()

    return text_html

def build_runs_html(runs):
    """Apply bold/italic formatting to each run."""
    runs_html = ""
    for run in runs:
        run_text = run.text.replace("<", "&lt;").replace(">", "&gt;")
        if getattr(run, "bold", False):
            run_text = f"<strong>{run_text}</strong>"
        if getattr(run, "italic", False):
            run_text = f"<em>{run_text}</em>"
        runs_html += run_text
    return runs_html

# ----------------------------------------------------------------
# FLASK MAIN
# ----------------------------------------------------------------
if __name__ == "__main__":
    # Run on port 5001 (change if needed)
    app.run(debug=True, host="0.0.0.0", port=5001)

