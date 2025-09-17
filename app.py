# app.py
from flask import Flask, render_template, request, flash, redirect, url_for, send_file
import os
from werkzeug.utils import secure_filename
import zipfile
import io
from generator_long import generate_certificates as gen_long
from generator_short import generate_certificates as gen_short

app = Flask(__name__)
app.secret_key = "siwes2025"
app.config['UPLOAD_FOLDER'] = 'uploads/excel'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/long", methods=["GET", "POST"])
def long_gen():
    if request.method == "POST":
        if 'excel' not in request.files:
            flash("❌ No file selected.")
            return redirect(request.url)
        file = request.files['excel']
        if file.filename == '':
            flash("❌ No file selected.")
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash("❌ Please upload a valid .xlsx file.")
            return redirect(request.url)

        filename = secure_filename(file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(excel_path)

        template_key = request.form.get('template')
        if template_key == 'template3':
            template_path = 'static/Template3.jpg'
        elif template_key == 'template4':
            template_path = 'static/Template4.jpg'
        elif template_key == 'template1':
            template_path = 'static/Template1.jpeg'
        elif template_key == 'template5':
            template_path = 'static/Template5.jpeg'
        elif template_key == 'template6':
            template_path = 'static/Template6.jpeg'
        else:
            flash("❌ Please select a template.")
            return redirect(request.url)

        # Generate certificates
        result = gen_long(excel_path, template_path)

        if result.get("error"):
            flash(f"❌ Error: {result['error']}", "error")
            return redirect(request.url)
        else:
            for err in result.get("errors", []):
                flash(f"⚠️ {err}", "warning")

        # Create ZIP in memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            output_dir = os.path.join(os.getcwd(), "generated_certificates")
            os.makedirs(output_dir, exist_ok=True)
            for f in os.listdir(output_dir):
                file_path = os.path.join(output_dir, f)
                if os.path.isfile(file_path):
                    zip_file.write(file_path, f)  # Add to ZIP
                    print(f"📎 Added to ZIP: {f}")

        zip_buffer.seek(0)

        # Auto-send ZIP
        generated_count = len(result.get("generated", []))
        zip_filename = f"certificates_{generated_count}.zip"
        flash(f"✅ Generated {generated_count} certificates!", "success")
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype="application/zip"
        )

    return render_template("long.html")


@app.route("/short", methods=["GET", "POST"])
def short_gen():
    if request.method == "POST":
        if 'excel' not in request.files:
            flash("❌ No file selected.")
            return redirect(request.url)
        file = request.files['excel']
        if file.filename == '':
            flash("❌ No file selected.")
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash("❌ Please upload a valid .xlsx file.")
            return redirect(request.url)

        filename = secure_filename(file.filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(excel_path)

        template_key = request.form.get('template')
        if template_key == 'template3':
            template_path = 'static/Template3.jpg'
        elif template_key == 'template4':
            template_path = 'static/Template4.jpg'
        elif template_key == 'template1':
            template_path = 'static/Template1.jpeg'
        elif template_key == 'template5':
            template_path = 'static/Template5.jpeg'
        elif template_key == 'template6':
            template_path = 'static/Template6.jpeg'
        else:
            flash("❌ Please select a template.")
            return redirect(request.url)

        result = gen_short(excel_path, template_path)

        if result.get("error"):
            flash(f"❌ Error: {result['error']}", "error")
            return redirect(request.url)
        else:
            for err in result.get("errors", []):
                flash(f"⚠️ {err}", "warning")

        # Create ZIP and auto-download
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            output_dir = os.path.join(os.getcwd(), "generated_certificates")
            os.makedirs(output_dir, exist_ok=True)
            for f in os.listdir(output_dir):
                file_path = os.path.join(output_dir, f)
                if os.path.isfile(file_path):
                    zip_file.write(file_path, f)
                    print(f"📎 Added to ZIP: {f}")

        zip_buffer.seek(0)

        generated_count = len(result.get("generated", []))
        zip_filename = f"certificates_{generated_count}.zip"
        flash(f"✅ Generated {generated_count} short certificates!", "success")
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype="application/zip"
        )

    return render_template("short.html")


@app.route("/download")
def download():
    output_dir = os.path.join(os.getcwd(), "generated_certificates")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    files = os.listdir(output_dir)
    return render_template("success.html", files=files)


@app.route("/download/<filename>")
def download_file(filename):
    output_dir = os.path.join(os.getcwd(), "generated_certificates")
    return send_file(os.path.join(output_dir, filename), as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)