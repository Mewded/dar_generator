from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
from dar_logic import generate_dar_summary

app = Flask(__name__)
app.secret_key = "g3tech_secret_key"  # needed for flashing messages

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def index():
    return render_template('dar_report.html')


@app.route('/generate', methods=['POST'])
def generate():
    # 1️⃣ Check for file upload
    if 'file' not in request.files:
        flash("No file part found in the request.")
        return redirect(url_for('index'))

    file = request.files['file']

    if file.filename == '':
        flash("No file selected.")
        return redirect(url_for('index'))

    # 2️⃣ Save the uploaded file
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    # 3️⃣ Run your DAR generator logic
    output_pdf = generate_dar_summary(filepath, OUTPUT_FOLDER)

    # 4️⃣ Send back the generated PDF file
    return send_file(output_pdf, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)



