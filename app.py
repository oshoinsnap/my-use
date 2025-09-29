import os
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
from werkzeug.utils import secure_filename
import tempfile
from email_name_merger import merge_by_email, read_list_from_excel, write_list_to_excel
from seprate import split_excel_by_industry
from cleaner import EmailListCleaner

app = Flask(__name__)

# For Vercel, use /tmp for temporary files
UPLOAD_FOLDER = '/tmp' if os.environ.get('VERCEL') else tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

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
        # Process the file
        output_filename = 'refined_' + filename
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        process_excel(filepath, output_filepath)
        return redirect(url_for('download_file', filename=output_filename))
    return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(filepath):
        response = send_file(filepath, as_attachment=True)
        # Optionally delete after download, but for simplicity, leave it
        return response
    return "File not found"

def process_excel(input_file, output_file):
    # Adapted from combine_excel.py
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

        # Get parameters
        source_sheet = request.form.get('source_sheet', 'Sheet1')
        target_sheet = request.form.get('target_sheet', 'Sheet2')
        output_filename = 'merged_' + filename
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        try:
            # Read sheets
            source_list = read_list_from_excel(filepath, source_sheet)
            target_list = read_list_from_excel(filepath, target_sheet)

            # Merge
            merged_list, matches = merge_by_email(source_list, target_list)

            # Write output
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

        # Get parameters
        industry_column = request.form.get('industry_column')
        output_format = request.form.get('output_format', 'separate_files')

        if not industry_column:
            flash('Industry column name is required')
            return redirect(url_for('index'))

        try:
            split_excel_by_industry(filepath, industry_column, output_format)
            flash('File split successfully. Check the industry_split_output folder.')
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

    # Allow CSV and Excel for cleaner
    allowed_exts = {'xlsx', 'xls', 'csv'}
    if '.' in file.filename and file.filename.rsplit('.', 1)[1].lower() in allowed_exts:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Get parameters
        email_column = request.form.get('email_column', 'email')
        advanced = request.form.get('advanced') == 'on'

        # Generate output filename
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

# For Vercel deployment
if __name__ == '__main__':
    app.run(debug=True)
