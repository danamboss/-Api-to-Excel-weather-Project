import os
import time
import random
from app import app
from flask import Flask, flash, request, redirect, render_template, url_for
from werkzeug.utils import secure_filename
from final import read_excel

ALLOWED_EXTENSIONS = set(['XLSX','xlsx','xls','XLS'])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    
@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/excel', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        #name = request.form["name"]
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected for uploading')
            time.sleep(2)
            return redirect(request.url)
        if file and allowed_file(file.filename):
            #save the name of the file
            filename = secure_filename(file.filename)
            #move file to upload folder
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            flash('File successfully uploaded')
            #return redirect('/')

            read_excel(os.path.join(app.config['UPLOAD_FOLDER'], filename))


            file.save(os.path.join(app.config['PROCESSED_FOLDER'], filename))
            flash('File successfully uploaded')
            
            print(os.path.join(app.config['PROCESSED_FOLDER'], filename))
            results = os.path.join(app.config['PROCESSED_FOLDER'], filename)
            return render_template('results.html', results=results)
        else:
            flash('That did not work, allowed file types are .txt')
            flash('Please try again')
            return redirect(request.url)

if __name__ == "__main__":  
    app.run( 
        host='0.0.0.0',
        port=random.randint(2000, 9000),  
        debug=True 
    )