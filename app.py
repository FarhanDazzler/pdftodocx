from flask import Flask, request, redirect, url_for, render_template, send_from_directory, send_file
from werkzeug.utils import secure_filename
import os 

from package.pdftoword import *

app = Flask(__name__)
# add 'uploads' folder to the project's root directory where uploaded files could be saved
UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024 # Limit the size of pdf file


ALLOWED_EXTENSIONS = {'pdf'}
def allowed_file(filename):
   return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
   if request.method == 'POST':
       if 'file' not in request.files:
           print('No file attached in request')
           return redirect(request.url)
       file = request.files['file']
       if file.filename == '':
           print('No file selected')
           return redirect(request.url)
       if file and allowed_file(file.filename):
           filename = secure_filename(file.filename)
           print("filename={}".format(os.path.abspath(filename)))
           file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
           print("Uploaded successfully")
           # my conversion code
           print(filename)
           res = pdfToDocx(filename=os.path.join(app.config['UPLOAD_FOLDER'], filename),
                     remote_name=filename,
                      output_name=filename)
           print("response from server", res)
           return render_template("download.html",output=filename.rsplit('.')[0]+'.docx')
        #    process_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), filename)
        #    return redirect(url_for('uploaded_file', filename=filename))
   return render_template('index.html')

@app.route('/download/<string:saved_file>')
def downloadFile(saved_file):
    # path = '3rd.docx'
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], saved_file, as_attachment=True)
    # return send_file(path, as_attachment=True)


@app.route('/uploads/<filename>')
def uploaded_file(filename):
   return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)