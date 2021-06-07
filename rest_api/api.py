import os
import flask
import pandas as pd
from flask import request, jsonify
from questions import python
from flask_cors import CORS
from werkzeug.utils import secure_filename

app = flask.Flask(__name__)
CORS(app)
app.config["DEBUG"] = True
app.config['UPLOAD_FOLDER'] = 'temp_csv'
ALLOWED_EXTENSIONS = set(['csv', 'xls', 'xlsx'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_csv_rows(filename):
    file_path = 'temp_csv/'+filename
    excel_file = pd.read_excel(file_path)
    return [x for x in excel_file.columns]


@app.route('/api/questions/python', methods=['GET'])
def api_all():
    args = request.args
    total_questions = int(args['totalquestions'])
    question_set = {k: python[k]
                    for k in sorted(python.keys())[:total_questions]}
    print(question_set)
    return jsonify(question_set), 200


@app.route('/api/get-excel-data', methods=['POST'])
def excel_rows():
    if 'file' not in request.files:
        resp = jsonify(
            {'message': 'No file part in the request', 'type': 'warning'})
        resp.status_code = 400
        return resp
    file = request.files['file']
    if file.filename == '':
        resp = jsonify(
            {'message': 'No file selected for uploading', 'type': 'warning'})
        resp.status_code = 400
        return resp
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        rows = get_csv_rows(filename)
        resp = jsonify({'rows': rows})
        resp.status_code = 201
        return resp
    else:
        resp = jsonify(
            {'message': 'Please upload a valid file.', 'type': 'error'})
        resp.status_code = 400
        return resp


app.run()
