import uuid, shutil, zipfile
from flask import render_template, send_from_directory
from app import app
from app.forms import FileForm
from app.excel import excel
from werkzeug.utils import secure_filename
from os import path, mkdir, chdir
from time import sleep

def allowed_file(filename):
    ALLOWED_EXTENSIONS = ['xlsx']
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    form = FileForm()
    if form.validate_on_submit():
        last_years_file = form.last_years_file.data
        lang = form.lang.data

        if last_years_file and allowed_file(last_years_file.filename):
            this_years_file_path = path.join(app.config['UPLOAD_FOLDER'], "5786.xlsx")
            output_directory = path.join(app.config['UPLOAD_FOLDER'], str(uuid.uuid4()).replace('-', ''))
            mkdir(output_directory)
            
            last_years_file_name = secure_filename(last_years_file.filename)

            last_years_file_path = path.join(output_directory, last_years_file_name)

            last_years_file.save(last_years_file_path)

            excel_return = excel(last_years_file_path, this_years_file_path, lang, output_directory)
            if excel_return == True:
                zip_path = f"{output_directory}.zip"
                chdir(app.config['UPLOAD_FOLDER'])
                zip_file = zipfile.ZipFile(f"{output_directory.rsplit('/', 1)[1]}.zip", 'w')
                zip_file.write(f"{output_directory.rsplit('/', 1)[1]}/output.xlsx", compress_type=zipfile.ZIP_DEFLATED)
                zip_file.write(f"{output_directory.rsplit('/', 1)[1]}/extras.xlsx", compress_type=zipfile.ZIP_DEFLATED)
                zip_file.close()
                return send_from_directory(app.config['UPLOAD_FOLDER'], zip_path.rsplit('/', 1)[1] , as_attachment=True)
            else:
                return excel_return
        return "No file or wrong filetype (must be .xlsx)!"

    return render_template('index.html', form=form)