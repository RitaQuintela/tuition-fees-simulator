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

    # Assuming the backend admin panel provides a way to upload the Excel file
    uploaded_file = request.files['excelFile']

    if uploaded_file and allowed_file(uploaded_file.filename):
        filename = secure_filename(uploaded_file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        uploaded_file.save(file_path)

        workbook = load_workbook(file_path)
        sheet = workbook.active

        credit_cell = 'A' + str(credits)
        tuition_cell = sheet[credit_cell]

        if tuition_cell.value is None:
            return 'Tuition data not found for the given number of credits.'

        tuition = tuition_cell.value

        return f'Tuition Fees for {credits} credits: ${tuition:.2f}'

    return 'Invalid file format or no file uploaded'

if __name__ == '__main__':
    app.run(debug=True)
