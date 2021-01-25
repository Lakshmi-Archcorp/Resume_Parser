import flask
from flask import Flask, render_template, jsonify, make_response, redirect, send_file
from flask import request
from flask_cors import CORS, cross_origin
from werkzeug.utils import secure_filename
from exceptionHandler import ExceptionHandler
from ResumeParser import *

app = Flask(__name__)
CORS(app)

@app.route('/home')
def query_example():
    return 'Todo...'

@app.route("/api/process", methods=['POST'])
def process():
    if flask.request.method == 'POST' :
        # uploaded_file = request.files['file']
        # if uploaded_file.filename != '':
        #     uploaded_file.save(uploaded_file.filename)
        file_Path = request.args.get('file_Path')
        # file_Path = "uploaded_file.filename"
        return json.dumps(extractResume(file_Path))

@app.errorhandler(Exception)
def handle_exception(e):
    return ExceptionHandler.handleError(e)

if __name__ == '__main__':
    app.run(debug=True)