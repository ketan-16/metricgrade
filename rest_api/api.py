import flask
from flask import request, jsonify
from questions import python
from flask_cors import CORS

app = flask.Flask(__name__)
CORS(app)
app.config["DEBUG"] = True


@app.route('/api/questions/python', methods=['GET'])
def api_all():
    args = request.args
    total_questions = int(args['totalquestions'])
    question_set = {k: python[k]
                    for k in sorted(python.keys())[:total_questions]}
    print(question_set)
    return jsonify(question_set), 200


app.run()
