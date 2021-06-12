import os
import flask
import pandas as pd
from flask import request, jsonify
from pandas.io import excel
from questions import python
from flask_cors import CORS
from werkzeug.utils import secure_filename
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import cred
import threading
import json

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


def get_col_data(filename):
    file_path = 'temp_csv/'+filename
    excel_file = pd.read_excel(file_path)
    return list(excel_file['Email'])


def validate(filename, stu_email, stu_prn):
    file_path = 'temp_csv/'+filename
    excel_file = pd.read_excel(file_path)
    excel_file = excel_file.filter(['PRN', 'Email'])
    tempdf = excel_file['PRN']
    for x in tempdf:
        if isinstance(x, int):
            stu_prn = int(stu_prn)
    excel_file = excel_file.loc[excel_file['PRN'] == stu_prn]
    excel_file = excel_file.loc[excel_file['Email'] == stu_email]

    if(len(excel_file)):
        return True
    return False


def send_test_mails(recipients, company_name, start_date, user_name, start_time):
    smtp_ssl_host = 'smtp.gmail.com'  # smtp.mail.yahoo.com
    smtp_ssl_port = 465
    username = cred.gmail[0]
    password = cred.gmail[1]
    sender = 'exam.metricgrade@gmail.com'
    targets = recipients  # ['ktnydv@gmail.com','test@gmail.com']

    msg = MIMEMultipart('alternative')
    msg['Subject'] = '{companyName} - Assesment Scheduled | {startDate}'.format(
        companyName=company_name, startDate=start_date)
    msg['From'] = sender
    msg['To'] = ', '.join(targets)
    text = ""
    html = """<!DOCTYPE html> <html lang="en"> <head> <meta charset="UTF-8"> <meta http-equiv="X-UA-Compatible" content="IE=edge"> <meta name="viewport" content="width=\, initial-scale=1.0"> <title>Document</title> </head> <body> <style>/* This styles you should add to your html as inline-styles */ /* You can easily do it with http://inlinestyler.torchboxapps.com/ */ /* Copy this html-window code converter and click convert button */ /* After that you can remove this style from your code */ /* This CSS code you should add to <head> of your page */ /* This code is for responsive design */ /* It didn't work in Gmail app on Android, but work fine on iOS */ @import url(https://fonts.googleapis.com/css?family=Roboto:400,700,400italic,700italic&subset=latin,cyrillic); @media only screen and (min-width: 0){.wrapper{text-rendering: optimizeLegibility;}}@media only screen and (max-width: 620px){[class=wrapper]{min-width: 302px !important; width: 100% !important;}[class=wrapper] .block{display: block !important;}[class=wrapper] .hide{display: none !important;}[class=wrapper] .top-panel, [class=wrapper] .header, [class=wrapper] .main, [class=wrapper] .footer{width: 302px !important;}[class=wrapper] .title, [class=wrapper] .subject, [class=wrapper] .signature, [class=wrapper] .subscription{display: block; float: left; width: 300px !important; text-align: center !important;}[class=wrapper] .signature{padding-bottom: 0 !important;}[class=wrapper] .subscription{padding-top: 0 !important;}}body{margin: 0; padding: 0; mso-line-height-rule: exactly; min-width: 100%;}.wrapper{display: table; table-layout: fixed; width: 100%; min-width: 620px; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;}body, .wrapper{background-color: #ffffff;}/* Basic */ table{border-collapse: collapse; border-spacing: 0;}table.center{margin: 0 auto; width: 602px;}td{padding: 0; vertical-align: top;}.spacer, .border{font-size: 1px; line-height: 1px;}.spacer{width: 100%; line-height: 16px}.border{background-color: #e0e0e0; width: 1px;}.padded{padding: 0 24px;}img{border: 0; -ms-interpolation-mode: bicubic;}.image{font-size: 12px;}.image img{display: block;}strong, .strong{font-weight: 700;}h1, h2, h3, p, ol, ul, li{margin-top: 0;}ol, ul, li{padding-left: 0;}a{text-decoration: none; color: #616161;}.btn{background-color: #2196F3; border: 1px solid #2196F3; border-radius: 2px; color: #ffffff; display: inline-block; font-family: Roboto, Helvetica, sans-serif; font-size: 14px; font-weight: 400; line-height: 36px; text-align: center; text-decoration: none; text-transform: uppercase; width: 200px; height: 36px; padding: 0 8px; margin: 0; outline: 0; outline-offset: 0; -webkit-text-size-adjust: none; mso-hide: all;}/* Top panel */ .title{text-align: left;}.subject{text-align: right;}.title, .subject{width: 300px; padding: 8px 0; color: #616161; font-family: Roboto, Helvetica, sans-serif; font-weight: 400; font-size: 12px; line-height: 14px;}/* Header */ .logo{padding: 16px 0;}/* Main */ .main{-webkit-box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.12), 0 1px 2px 0 rgba(0, 0, 0, 0.24); -moz-box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.12), 0 1px 2px 0 rgba(0, 0, 0, 0.24); box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.12), 0 1px 2px 0 rgba(0, 0, 0, 0.24);}/* Content */ .columns{margin: 0 auto; width: 600px; background-color: #ffffff; font-size: 14px;}.column{text-align: left; background-color: #ffffff; font-size: 14px;}.column-top{font-size: 24px; line-height: 24px;}.content{width: 100%;}.column-bottom{font-size: 8px; line-height: 8px;}.content h1{margin-top: 0; margin-bottom: 16px; color: #212121; font-family: Roboto, Helvetica, sans-serif; font-weight: 400; font-size: 20px; line-height: 28px;}.content p{margin-top: 0; margin-bottom: 16px; color: #212121; font-family: Roboto, Helvetica, sans-serif; font-weight: 400; font-size: 16px; line-height: 24px;}.content .caption{color: #616161; font-size: 12px; line-height: 20px;}/* Footer */ .signature, .subscription{vertical-align: bottom; width: 300px; padding-top: 8px; margin-bottom: 16px;}.signature{text-align: left;}.subscription{text-align: right;}.signature p, .subscription p{margin-top: 0; margin-bottom: 8px; color: #616161; font-family: Roboto, Helvetica, sans-serif; font-weight: 400; font-size: 12px; line-height: 18px;}</style> <center class="wrapper"> <table class="top-panel center" width="602" border="0" cellspacing="0" cellpadding="0"> <tbody> <tr> <td class="title" width="300">MetricGrade</td><td class="subject" width="300"><a class="strong" href="#" target="_blank">www.metricgrade.io</a> </td></tr><tr> <td class="border" colspan="2">&nbsp;</td></tr></tbody> </table> <div class="spacer">&nbsp;</div><table class="main center" width="602" border="0" cellspacing="0" cellpadding="0"> <tbody> <tr> <td class="column"> <div class="column-top">&nbsp;</div><table class="content" border="0" cellspacing="0" cellpadding="0"> <tbody> <tr> <td class="padded"> <h1>""" + \
        company_name+""" - Test Scheduled | """+start_date+"""</h1> <p>Greetings, """+user_name+""". We hope you're doing well. Your Assesment has been scheduled from<strong> """+start_date+""", """ + start_time + \
        """ Onwards</strong> by MetricGrade.</p><p>Ensure to appear for the test as per your convenience.</p><p>Please take note of the test rules:</p><p>1. Ensure that test is completed within alotted window.</p><p>2. Test is <b>Camera Proctored. </b> Any form of malpractice(s), if found, shall be acted upon with strict rules.</p><p>3. Do not leave the screen after the test Starts. Doing so may lead to automatic closing of test.</p><br><br><p>We greet you Best wishes for you test! In case of any problem, feel free to connect with our support team for quick and easy resolutions:</p><p style="text-align:center;"><a href="mailto:exam.metricgrade@gmail.com" class="btn">Get in Touch</a></p><p style="text-align:center;"> <a href="#" class="strong">Visit Site</a> </p><p class="caption">This is an auto-generated email.</p></td></tr></tbody> </table> <div class="column-bottom">&nbsp;</div></td></tr></tbody> </table> <div class="spacer">&nbsp;</div><table class="footer center" width="602" border="0" cellspacing="0" cellpadding="0"> <tbody> <tr> <td class="border" colspan="2">&nbsp;</td></tr><tr> <td class="signature" width="300"> <p> With best regards,<br>MetricGrade.io<br>+91 7040335652, Ketan Y<br></p><p> Support: <a class="strong" href="mailto:#" target="_blank">exam.metricgrade@gmail.com</a> </p></td><td class="subscription" width="300"> <div class="logo-image"> <a href="https://zavoloklom.github.io/material-design-iconic-font/" target="_blank"><img src="https://i.ibb.co/Sct3Rxj/nav-brand.png" alt="logo-alt" height="28px" style="margin-bottom:10px"></a> </div><p> <a class="strong block" href="#" target="_blank"> Unsubscribe </a> <span class="hide">&nbsp;&nbsp;|&nbsp;&nbsp;</span> <a class="strong block" href="#" target="_blank"> Account Settings </a> </p></td></tr></tbody> </table> </center> </body> </html>"""

    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    msg.attach(part1)
    msg.attach(part2)

    server = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
    server.login(username, password)
    server.sendmail(sender, targets, msg.as_string())
    server.quit()


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
        file.save(os.path.join(
            app.config['UPLOAD_FOLDER'], 'filetorename.xls'))
        rows = get_csv_rows('filetorename.xls')
        resp = jsonify({'rows': rows})
        resp.status_code = 201
        return resp
    else:
        resp = jsonify(
            {'message': 'Please upload a valid file.', 'type': 'error'})
        resp.status_code = 400
        return resp


@app.route('/api/setfilename', methods=['GET'])
def set_filename():
    args = request.args
    dname = str(args['drivename']+'.xls')
    src = os.path.join(
        app.config['UPLOAD_FOLDER'], 'filetorename.xls')
    dest = os.path.join(
        app.config['UPLOAD_FOLDER'], dname)
    try:
        os.rename(src, dest)
    except:
        pass
    resp = jsonify(
        {'message': 'File Renamed successfully'})
    resp.status_code = 200
    return resp


@app.route('/api/validatestudent', methods=['GET'])
def validate_student():
    args = request.args
    stu_email = args['email']
    filename = args['drivename']
    stu_prn = args['prn']
    to_show = validate(filename+'.xls', stu_email, stu_prn)
    resp = jsonify({'show': to_show})
    resp.status_code = 200
    return resp


@app.route('/api/sendstudentemail', methods=['POST'])
def send_test_emails():
    record = json.loads(request.data)
    print(record.to_dict())
    # email_recipients = get_col_data(filename+'.xls')
    # company_name = filename.split('_')[0]
    # start_date = start_date
    # user_name = display_name
    # start_time = start_time
    # send_email_thread = threading.Thread(target=sent_test_mails, args=(
    #     email_recipients, company_name, start_date, user_name, start_time,))
    # send_email_thread.start()
    # print('Email Job Started...')


app.run()
