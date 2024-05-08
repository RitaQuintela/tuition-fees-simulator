from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
import os
from openpyxl import load_workbook

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/calculate', methods=['POST'])
def calculate():
    credits = int(request.form['credits'])

    if 'excelFile' not in request.files:
        return 'No file part'

    file = request.files['excelFile']

    if file.filename == '':
        return 'No selected file'

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        workbook = load_workbook(file_path)
        sheet = workbook.active

        credit_cell = 'A' + str(credits)
        tuition_cell = sheet[credit_cell]

        if tuition_cell.value is None:
            return 'Tuition data not found for the given number of credits.'

        tuition = tuition_cell.value

        return f'Tuition Fees for {credits} credits: ${tuition:.2f}'

    return 'Invalid file format'

if __name__ == '__main__':
    app.run(debug=True)
