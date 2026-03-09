# app.py
from flask import Flask, render_template, request, flash, send_file, session, redirect
import os
from werkzeug.utils import secure_filename
import zipfile
import io

app = Flask(__name__)
app.secret_key = "siwes2025"  # Needed for sessions

# -----------------------------
# Hardcoded Admin & Users DB
# -----------------------------
USERS = {
    "admin": "superkey"  # Admin with your requested password
}

app.config['UPLOAD_FOLDER'] = 'uploads/excel'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# 🔐 Protect all routes except login
@app.before_request
def require_login():
    if 'username' not in session and request.endpoint not in ('login', 'health'):
        return redirect('/login')


# 🖥️ Main Certificate Page (formerly /long)
@app.route("/", methods=["GET", "POST"])
def index():
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
        if template_key == 'template5':
            template_path = 'static/Template5.jpeg'
        elif template_key == 'template6':
            template_path = 'static/Template6.jpeg'
        elif template_key == 'template7':
            template_path = 'static/Template7.jpg'
        elif template_key == 'template8':
            template_path = 'static/Template8.jpg'
        else:
            flash("❌ Please select a template.")
            return redirect(request.url)

        from generator_long import generate_certificates
        result = generate_certificates(excel_path, template_path)

        if result.get("error"):
            flash(f"❌ Error: {result['error']}", "error")
        else:
            for err in result.get("errors", []):
                flash(f"⚠️ {err}", "warning")

        # Create ZIP to download
        zip_buffer = io.BytesIO()
        output_dir = os.path.join(os.getcwd(), "generated_certificates")
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f in os.listdir(output_dir):
                file_path = os.path.join(output_dir, f)
                if os.path.isfile(file_path):
                    zipf.write(file_path, f)
        zip_buffer.seek(0)

        count = len(result.get("generated", []))
        flash(f"✅ Generated {count} certificates!", "success")
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f"certificates_{count}.zip",
            mimetype="application/zip"
        )

    return render_template("long.html")

@app.route("/short", methods=["GET", "POST"])
def short_certificates():
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

        if template_key == 'template5':
            template_path = 'static/Template5.jpeg'
        elif template_key == 'template6':
            template_path = 'static/Template6.jpeg'
        elif template_key == 'template7':
            template_path = 'static/Template7.jpg'
        elif template_key == 'template8':
            template_path = 'static/Template8.jpg'
        else:
            flash("❌ Please select a template.")
            return redirect(request.url)

        from generator_short import generate_certificates
        result = generate_certificates(excel_path, template_path)

        if result.get("error"):
            flash(f"❌ Error: {result['error']}", "error")

        # ZIP results
        zip_buffer = io.BytesIO()
        output_dir = os.path.join(os.getcwd(), "generated_certificates")

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f in os.listdir(output_dir):
                file_path = os.path.join(output_dir, f)
                if os.path.isfile(file_path):
                    zipf.write(file_path, f)

        zip_buffer.seek(0)

        count = len(result.get("generated", []))
        flash(f"✅ Generated {count} certificates!", "success")

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f"certificates_{count}.zip",
            mimetype="application/zip"
        )

    return render_template("short.html")


# 🔐 Login Route
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        if username in USERS and USERS[username] == password:
            session['username'] = username
            flash("✅ Logged in successfully!", "success")
            return redirect("/")
        else:
            flash("❌ Invalid credentials.", "error")
    return render_template("login.html")


# 🧑‍💼 Register New User (Admin Only)
@app.route("/register", methods=["GET", "POST"])
def register():
    if session.get("username") != "admin":
        flash("⚠️ Admin access required.", "warning")
        return redirect("/")

    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")
        if not username or not password:
            flash("❌ Fill all fields.", "error")
        elif username in USERS:
            flash(f"⚠️ User '{username}' already exists.", "warning")
        else:
            USERS[username] = password
            flash(f"✅ User '{username}' created!", "success")
            return redirect("/register")

    return render_template("register.html", users=USERS)


# 🚪 Logout
@app.route("/logout")
def logout():
    username = session.pop("username", None)
    flash(f"👋 {username}, logged out.", "info")
    return redirect("/login")

# 🩺 Health Check Endpoint
@app.route("/health")
def health():
    return {"status": "ok"}, 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)