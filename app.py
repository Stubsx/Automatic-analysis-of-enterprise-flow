from csv import excel
import os
import time
from tkinter.tix import Tree
from turtle import stamp

from flask import Flask, render_template, request, redirect, url_for, send_from_directory,flash
from werkzeug.utils import secure_filename

from function.bankvis import workflow
# Initialize the Flask application
app = Flask(__name__)
stamp = 0
excel_output = False

# This is the path to the upload directory
app.config['UPLOAD_FOLDER'] = 'upload/'

# These are the extension that we are accepting to be uploaded
app.config['ALLOWED_EXTENSIONS'] = set(['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif','xlsx','xls'])

# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
  return '.' in filename and \
      filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']
      
# This route will show a form to perform an AJAX request jQuery is loaded to execute the request and update the value of the operation
@app.route('/')
def index():
  return render_template('index.html')

# Route that will process the file upload
@app.route('/upload', methods=['POST'])
def upload():
    # Get the name of the uploaded files
    uploaded_files = request.files.getlist("file[]")
    
    global stamp
    stamp = time.time()
    # Change time format to yyyy-mm-dd
    stamp = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime(stamp))
    excelbool = request.form.get('excelbool')
    if excelbool=='on':
        excelbool=True
    else:
        excelbool=False
    
    global excel_output
    excel_output = excelbool


    os.makedirs(app.config['UPLOAD_FOLDER']+stamp+'/')
    filenames = []
    for file in uploaded_files:
        # Check if the file is one of the allowed types/extensions
        if file and allowed_file(file.filename):       
            # Make the filename safe, remove unsupported chars
            # filename = secure_filename(file.filename)
            filename  = file.filename
            
            # Move the file form the temporal folder to the upload folder we setup
            file.save(os.path.join(app.config['UPLOAD_FOLDER']+stamp+'/',filename))
            
            # Save the filename into a list, we'll use it later
            filenames.append(filename)
            
    from function import bankvis
    df = bankvis.mergeflow('./upload/'+stamp)
    workflow(df,output_excel=excelbool)

    return render_template('upload.html', filenames=filenames,excel_output=excel_output)

# This route is expecting a parameter containing the name
# of a file. Then it will locate that file on the upload
# directory and show it on the browser, so if the user uploads
# an image, that image is going to be show after the upload
@app.route('/upload/<filename>')
def uploaded_file(filename):
    global stamp
    return send_from_directory(app.config['UPLOAD_FOLDER']+stamp+'/',
                    filename)

@app.route('/output/<filename>')
def output_file(filename):
    return send_from_directory('output/',filename)
  
@app.route('/flowvisual')  
def flowvisual():
    return render_template('流水可视化分析.html')

if __name__ == '__main__':
  app.run(debug=True)
