import os
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
import tempfile
from email_name_merger import merge_by_email, read_list_from_excel, write_list_to_excel
from seprate import split_excel_by_industry
from cleaner import EmailListCleaner

app = Flask(__name__, static_url_path='/static', static_folder='static')
app.secret_key = 'your-secret-key-change-this-in-production'

# For Vercel, use /tmp for temporary files
UPLOAD_FOLDER = '/tmp' if os.environ.get('VERCEL') else tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/index.html')
def index():
    return render_template('index.html')

@app.route('/')
def root_redirect():
    return redirect(url_for('index'))

@app.route('/combiner.html')
def combiner():
    return render_template('combiner.html')

@app.route('/merger.html')
def merger():
    return render_template('merger.html')

@app.route('/splitter.html')
def splitter():
    return render_template('splitter.html')

@app.route('/cleaner.html')
def cleaner():
    return render_template('cleaner.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        output_filename = 'refined_' + filename
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        try:
            process_excel(filepath, output_filepath)
            return redirect(url_for('download_file', filename=output_filename))
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
    return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found"

def process_excel(input_file, output_file):
    sheets = pd.read_excel(input_file, sheet_name=None)
    combined = pd.concat(sheets.values(), ignore_index=True)
    possible_email_cols = ['email', 'Email', 'email address', 'Email Address', 'e-mail', 'E-mail']
    email_col = None
    for col in combined.columns:
        if col.lower() in [p.lower() for p in possible_email_cols]:
            email_col = col
            break
    if email_col is None:
        raise ValueError("No email column found.")
    unique = combined.drop_duplicates(subset=email_col)
    unique.to_excel(output_file, index=False)

@app.route('/merge_names', methods=['POST'])
def merge_names():
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        source_sheet = request.form.get('source_sheet', 'Sheet1')
        target_sheet = request.form.get('target_sheet', 'Sheet2')
        output_filename = 'merged_' + filename
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        try:
            source_list = read_list_from_excel(filepath, source_sheet)
            target_list = read_list_from_excel(filepath, target_sheet)
            merged_list, matches = merge_by_email(source_list, target_list)
            success = write_list_to_excel(merged_list, output_filepath)
            if success:
                flash(f'Successfully merged {matches} records')
                return redirect(url_for('download_file', filename=output_filename))
            else:
                flash('Error writing merged file')
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
    return redirect(url_for('index'))

@app.route('/split_industry', methods=['POST'])
def split_industry():
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        industry_column = request.form.get('industry_column')
        output_format = request.form.get('output_format', 'separate_files')

        if not industry_column:
            flash('Industry column name is required')
            return redirect(url_for('index'))

        try:
            base_name = os.path.splitext(filename)[0]
            if output_format == 'single_file_multiple_sheets':
                output_filename = f'{base_name}_split_by_{industry_column}.xlsx'
                output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                split_excel_by_industry(filepath, industry_column, output_format, output_filepath, verbose=False)
                flash('File split successfully into multiple sheets.')
                return redirect(url_for('download_file', filename=output_filename))
            else:
                import zipfile, shutil
                zip_filename = f'{base_name}_split_by_{industry_column}.zip'
                zip_filepath = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
                temp_split_dir = os.path.join(app.config['UPLOAD_FOLDER'], f'temp_split_{base_name}')
                os.makedirs(temp_split_dir, exist_ok=True)
                split_excel_by_industry(filepath, industry_column, output_format, temp_split_dir, verbose=False)
                with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, _, files in os.walk(temp_split_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_split_dir)
                            zipf.write(file_path, arcname)
                shutil.rmtree(temp_split_dir)
                flash('File split successfully into separate files (ZIP archive).')
                return redirect(url_for('download_file', filename=zip_filename))
        except Exception as e:
            flash(f'Error splitting file: {str(e)}')
    return redirect(url_for('index'))

@app.route('/clean_emails', methods=['POST'])
def clean_emails():
    if 'file' not in request.files:
        flash('No file uploaded')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))

    allowed_exts = {'xlsx', 'xls', 'csv'}
    if '.' in file.filename and file.filename.rsplit('.', 1)[1].lower() in allowed_exts:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        email_column = request.form.get('email_column', 'email')
        advanced = request.form.get('advanced') == 'on'
        input_path = os.path.splitext(filename)
        output_filename = f"{input_path[0]}_cleaned{input_path[1]}"
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        try:
            cleaner = EmailListCleaner()
            cleaned_df = cleaner.clean_email_list(filepath, email_column, output_filepath, advanced)
            if cleaned_df is not None:
                flash('Email list cleaned successfully')
                return redirect(url_for('download_file', filename=output_filename))
            else:
                flash('Error cleaning email list')
        except Exception as e:
            flash(f'Error cleaning emails: {str(e)}')
    else:
        flash('Invalid file format. Use CSV or Excel.')
    return redirect(url_for('index'))
