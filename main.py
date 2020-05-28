import os, glob, openpyxl as xl
from flask import Flask, flash, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

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
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            else:
                flash('Unsupported File type: Only .xls and .xlsx files are allowed.')
                return redirect(request.url)
        if len(filelist) == 1:
            flash('How do you expect us to merge one file?')
            return redirect(request.url)
        merge(filelist)
        return redirect(url_for('index',
                                filename='output.xlsx'))
    return render_template('index.html')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def merge(filelist):
    out = xl.Workbook()
    out_sheet = out.worksheets[0]

    for file in glob.glob(UPLOAD_FOLDER + '\\*.xlsx'):
        ws = xl.load_workbook(file).worksheets[0]
        temp_row = []
        for row in ws:
            for cell in row:
                temp_row.append(cell.value)
            out_sheet.append(temp_row)

    out.save(UPLOAD_FOLDER + '\\output.xlsx')

if __name__ == '__main__':
    app.secret_key = 'super secret key'
    app.config['SESSION_TYPE'] = 'filesystem'

    app.debug = True
    app.run()
