import os, glob, openpyxl as xl, random, string
from flask import Flask, flash, request, redirect, url_for, render_template, send_from_directory
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'tmp/uploads'
OUTPUT_FOLDER = 'merged'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'None')

@app.route('/', methods=['GET', 'POST'])
def index(error=''):
    if request.method == 'POST':
        filelist = []
        if 'files' not in request.files:
            flash('File Not Found: Attach atleast one file.')
            return redirect(request.url)
        files = request.files.getlist('files')
        for file in files:
            if file.filename == '':
                flash('File Not Found: Attach atleast one file.')
                return redirect(request.url)
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filelist.append(filename)
                file.save(os.path.join(UPLOAD_FOLDER, filename))
            else:
                flash('Unsupported File type: Only .xls and .xlsx files are allowed.')
                return redirect(request.url)
        if len(filelist) == 1:
            flash('How do you expect us to merge one file?')
            return redirect(request.url)
        return render_template('index.html', filename = merge(filelist))
    return render_template('index.html')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def merge(filelist):
    out = xl.Workbook()
    out_sheet = out.worksheets[0]
    output = OUTPUT_FOLDER + '\\merged-' + ''.join(random.choices(string.ascii_lowercase + string.digits, k=10)) + '.xlsx'
    for file in glob.glob(UPLOAD_FOLDER + '\\*.xlsx'):
        ws = xl.load_workbook(file).worksheets[0]
        temp_row = []
        for row in ws:
            for cell in row:
                temp_row.append(cell.value)
            out_sheet.append(temp_row)

    out.save(output)
    return output

@app.route('/merged/<filename>')
def download(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run()
