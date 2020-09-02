from flask import Flask, request, redirect, url_for, render_template, send_from_directory, send_file
# from pdftoword import *
import os
import sys
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException

#### for pdftoword #####################
import groupdocs_conversion_cloud
from shutil import copyfile
import os
from PyPDF2 import PdfFileReader
#################################

# groupdocx API settings
api_sid = os.environ.get('APP_SID')
api_key = os.environ.get('APP_KEY')


#### function to convert pdftoword
def pdfToDocx(filename, remote_name, output_name, docfor):
        app_sid = api_sid
        app_key = api_key
        print("KEYS accepted")
        # no. of pages in pdf file
        print(application.config['UPLOAD_FOLDER']+remote_name)
        if docfor == 'docx':
            file  = PdfFileReader(open(application.config['UPLOAD_FOLDER']+remote_name,'rb'))
            page_counts = file.getNumPages()
        elif docfor == 'pdf':
            page_counts = int(999)
        print("Number of pages: ", page_counts)
        print("Download path: ", application.config['DOWNLOAD_FOLDER']+output_name)
        # Create instance of the API
        convert_api = groupdocs_conversion_cloud.ConvertApi.from_keys(app_sid, app_key)
        file_api = groupdocs_conversion_cloud.FileApi.from_keys(app_sid, app_key)

        try:

                #upload soruce file to storage
                filename = filename
                remote_name = remote_name
                output_name= remote_name.rsplit('.')[0]+'.'+docfor
                strformat=docfor
                print(filename, remote_name, output_name, docfor)

                request_upload = groupdocs_conversion_cloud.UploadFileRequest(remote_name,filename)
                print("upload request")
                response_upload = file_api.upload_file(request_upload)
                print("uploaded to the cloud")
                        #Convert PDF to Word document
                settings = groupdocs_conversion_cloud.ConvertSettings()
                settings.file_path =remote_name
                settings.format = strformat
                settings.output_path = output_name
                print("document converted")

                loadOptions = groupdocs_conversion_cloud.PdfLoadOptions()
                loadOptions.hide_pdf_annotations = True
                loadOptions.remove_embedded_files = False
                loadOptions.flatten_all_fields = True

                settings.load_options = loadOptions

                convertOptions = groupdocs_conversion_cloud.DocxConvertOptions()
                convertOptions.from_page = 1
                convertOptions.pages_count = int(page_counts)

                settings.convert_options = convertOptions

                request = groupdocs_conversion_cloud.ConvertDocumentRequest(settings)
                response = convert_api.convert_document(request)
                
                print("Document converted successfully: " + str(response))
                # Download request
                request_download = groupdocs_conversion_cloud.DownloadFileRequest(output_name)
                response_download = file_api.download_file(request_download)
                print("Response Download ", response_download)
                copyfile(response_download,
                        application.config['DOWNLOAD_FOLDER']+output_name)
                
                print("Successful")
                return output_name
        except groupdocs_conversion_cloud.ApiException as e:
                print("Exception when calling get_supported_conversion_types: {0}".format(e.message))
                return render_template("api_exception.html")

######## function for docx to pdf #############################
def docxToPdf():
    pass

#################################################

### Flask instance named application
application = Flask(__name__)
UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
application.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'
application.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

application.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024 # Limit the size of pdf file

## allowed file extension function
ALLOWED_EXTENSIONS = {'pdf', 'docx'}
def allowed_file(filename):
   return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# application.config.from_object(os.environ.get('APP_SETTINGS'))


@application.route('/', methods=['GET', 'POST'])
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
           file.save(os.path.join(application.config['UPLOAD_FOLDER'], filename))
           print("Uploaded successfully")
           # my conversion code
           print(filename)
           res = pdfToDocx(filename=os.path.join(application.config['UPLOAD_FOLDER'], filename),
                     remote_name=filename,
                      output_name=filename,
                      docfor='docx')
           print("response from server", res)
           return render_template("download2.html",output=res) #res is alternative of: filename.rsplit('.')[0]+'.docx' 
   return render_template('index.html')

@application.route('/doctopdf', methods=['GET','POST'])
def toPdf():
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
           file.save(os.path.join(application.config['UPLOAD_FOLDER'], filename))
           print("Uploaded successfully")
           # my conversion code
           print(filename)
           res = pdfToDocx(filename=os.path.join(application.config['UPLOAD_FOLDER'], filename),
                     remote_name=filename,
                      output_name=filename,
                      docfor='pdf')
           print("response from server", res)
           return render_template("download2.html",output=res)
    return render_template("doctopdf.html")

@application.route('/download/<string:saved_file>')
def downloadFile(saved_file):
    return send_from_directory(application.config['DOWNLOAD_FOLDER'], saved_file, as_attachment=True)
    


######################################### EXCEPTION HANDLING ##############################################


##############################  Handling HTTP Exceptions
@application.errorhandler(Exception)
def handle_exception(e):
    # pass through HTTP errors
    if isinstance(e, HTTPException):
        return e

    # now you're handling non-HTTP exceptions only
    return render_template("unhandlederror.html", e=e), 500


if  __name__ == '__main__':
        application.run()
