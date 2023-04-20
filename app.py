import os
import pandas as pd
from flask import Flask, request, redirect, url_for, flash, send_file
from werkzeug.utils import secure_filename


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.secret_key = 'supersecretkey'

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def get_sql_data_type(column_data):
    is_int = all(column_data.dropna().apply(lambda x: isinstance(x, int)))
    is_float = all(column_data.dropna().apply(lambda x: isinstance(x, (float, int))))
    is_date = all(column_data.dropna().apply(lambda x: isinstance(x, pd.Timestamp)))
    
    if is_int:
        return 'INT'
    elif is_float:
        return 'FLOAT'
    elif is_date:
        return 'DATETIME'
    else:
        max_len = column_data.astype(str).apply(len).max()
        return f'NVARCHAR({max_len})'

def generate_sql_code(df, table_name):
    column_types = []

    for column in df.columns:
        column_data = df[column]
        column_type = f'[{column}] {get_sql_data_type(column_data)}'
        column_types.append(column_type)

    create_table = f"CREATE TABLE [{table_name}] (\n" + ",\n".join(column_types) + "\n);\n\n"

    insert_values = []
    for index, row in df.iterrows():
        values = []
        for value in row:
            if pd.isna(value):
                values.append('NULL')
            elif isinstance(value, (int, float)):
                values.append(str(value))
            elif isinstance(value, pd.Timestamp):
                values.append(f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'")
            else:
                escaped_value = repr(value).replace("'", "'")[1:-1]
                values.append(f"N'{escaped_value}'")

        insert_values.append(f"INSERT INTO [{table_name}] VALUES (" + ', '.join(values) + ");")

    return create_table + "\n".join(insert_values)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            table_name = request.form['table_name']
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

            df = pd.read_excel(file_path)
            sql_code = generate_sql_code(df, table_name)

            with open(f'{table_name}.sql', 'w', encoding='utf-8') as f:
                f.write(sql_code)

            os.remove(file_path)

            return send_file(f'{table_name}.sql', as_attachment=True, attachment_filename=f'{table_name}.sql')

    return '''
    <!doctype html>
    <title>Excel to SQL Converter</title>
    <h1>Upload Excel file to convert to SQL Server</h1>
    <form method=post enctype=multipart/form-data>
      <input type=text name=table_name placeholder="Table name">
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

if __name__ == '__main__':
    app.run(debug=True)
