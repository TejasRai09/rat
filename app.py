from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename
import os
import logic_combined as logic

UPLOAD_DIR   = "uploads"
ALLOWED_EXT  = {".xlsx", ".xls"}

app = Flask(__name__)
app.secret_key = "replace-me"
app.config["UPLOAD_FOLDER"] = UPLOAD_DIR
os.makedirs(UPLOAD_DIR, exist_ok=True)


# def _save(f):
#     name = secure_filename(f.filename)
#     path = os.path.join(UPLOAD_DIR, name)
#     f.save(path)
#     return path
def _save(f):
    name = secure_filename(f.filename)
    base, ext = os.path.splitext(name)

    counter = 1
    path = os.path.join(UPLOAD_DIR, name)

    # Ensure unique filename
    while os.path.exists(path):
        path = os.path.join(UPLOAD_DIR, f"{base}_{counter}{ext}")
        counter += 1

    f.save(path)
    return path

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        kotak   = request.files.get("kotak")
        est     = request.files.get("estimation")
        section = request.form.get("section") or "both"

        if not kotak or not est:
            flash("Upload both files first!", "danger")
            return redirect(url_for("index"))

        for f in (kotak, est):
            if os.path.splitext(f.filename)[1].lower() not in ALLOWED_EXT:
                flash(f"File type not supported: {f.filename}", "danger")
                return redirect(url_for("index"))

        kotak_path = _save(kotak)
        est_path   = _save(est)

        try:
            out_path = logic.update(kotak_path, est_path, section)
            flash(f"{section.title()} section updated successfully!", "success")
            return send_file(out_path, as_attachment=True)
        except Exception as e:
            flash(f"Processing error: {e}", "danger")
            return redirect(url_for("index"))

    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
