from flask import Flask

UPLOAD_FOLDER = 'uploads/'
PROCESSED_FOLDER = 'static/processed/'

app = Flask(  
	__name__,
	template_folder='templates',  # Name of html file folder
	static_folder='static'  # Name of directory for static files
)
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 3 * 1024 * 1024