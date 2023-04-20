import os
import pandas as pd
import pyodbc
from flask import Flask, request, redirect, flash, send_file, url_for
from werkzeug.utils import secure_filename
from os.path import join

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists('downloads'):
    os.makedirs('downloads')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_excel_to_sql(file_path, table_name):
    df = pd.read_excel(file_path, engine='openpyxl')
    # ... (resto del código para convertir a SQL y guardar en "downloads/output.sql") ...

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            table_name = request.form['table_name']
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            convert_excel_to_sql(os.path.join(app.config['UPLOAD_FOLDER'], filename), table_name)
            return redirect(url_for('download_file_page'))

    return '''
    <!doctype html>
    <title>Upload Excel file</title>
    <h1>Upload Excel file</h1>
    <form method=post enctype=multipart/form-data>
      <label for="table_name">Table name:</label>
      <input type="text" id="table_name" name="table_name" required><br><br>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

@app.route('/download')
def download_file_page():
    return '''
    <!doctype html>
    <title>Download SQL file</title>
    <h1>Download SQL file</h1>
    <a href="/downloads/output.sql">Click here to download</a>
    '''

@app.route('/downloads/<name>')
def download_file(name):
    return send_file(os.path.join("downloads", name), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
