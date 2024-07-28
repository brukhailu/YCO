from time import strftime
import random
import mysql.connector
from flask import Flask, render_template, flash, request, send_file, jsonify,url_for
from werkzeug.utils import secure_filename
from wtforms import Form, StringField
from wtforms.validators import DataRequired
from svb import *

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
if not os.path.exists("uploads"):
    os.makedirs(os.path.join(os.getcwd(), 'uploads'))
UPLOAD_FOLDER = 'uploads/'
DEBUG = True
app = Flask(__name__)
app.config.from_object(__name__)
app.config['SECRET_KEY'] = 'SjdnUends821Jsdlkvxh391ksdODnBjdDw'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

class ReusableForm(Form):
    rrn_column = StringField('rrn_column:', validators=[DataRequired()])
    amount_column = StringField('amount_column:', validators=[DataRequired()])

def get_time():
    time = strftime("%Y%m%d%H%M")
    return time

def write_to_disk(filename, rrn_column, pan_column , amount_column):
    data = open('file.log', 'a')
    timestamp = get_time()
    data.write('DateStamp = {}, File Name = {}, RRN Column = {}, PAN Column = {}, Amount Column = {} \n'.format(timestamp, filename, rrn_column, pan_column, amount_column))
    data.close()
global output_file_name

@app.route('/<variable>')
def download_file(variable):
    try:
        path = os.getcwd() + "/" + variable + "/RESULT."+variable+".xls"
        return send_file(path, download_name = str("RESULT."+variable+".xls"), as_attachment=True)
    except:
        return render_template('error.html', error = 'Download Error!!')

@app.route('/svb', methods=['GET', 'POST'])
def svb():
    form1 = ReusableForm(request.form)
    if request.method == 'POST':
        if form1.validate():
            bank = request.form.get('bank')
            save = request.form.get('save')
            atm = request.form.get('atm')
            hst = request.form.get('hst')
            fix = request.form.get('fix')
            institution = request.form.get('institution')
            rrn_column = (int(request.form['rrn_column'])) - 1
            pan_column = (int(request.form['pan_column'])) - 1
            amount_column = (int(request.form['amount_column'])) - 1
            file = request.files['file']

            if bank == None or institution == None or file.filename == '':
                flash('Error: All Fields are Required')
            else:
                global filename
                output_file_name = str(get_time() + "." + institution + "." + bank + "." + str(random.randint(100000, 999999)))
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                path = str(os.path.join(UPLOAD_FOLDER, str(filename)))
                # if institution == 'ET-Switch':
                # try:
                main(atm,save,hst,fix,institution,output_file_name,bank, path, rrn_column, pan_column, amount_column)
                write_to_disk(filename, rrn_column, pan_column, amount_column)
                flash('{}, [RRN Column = {}, PAN Column = {}, Amount Column = {}]'.format(filename, rrn_column,pan_column, amount_column))
                return render_template('svb_result.html', rrn_column=rrn_column + 1, pan_column=pan_column + 1, amount_column=amount_column + 1,bank=bank,output_file_name=output_file_name,filename=filename, institution=institution, save=save)
                # except:
                #     return render_template('error.html', error = 'OPS, Error Occurred, Contact Your Admin!')
        else:
            flash('Error: All Fields are Required')
    return render_template('svb.html', form1=form1)

@app.route('/fix', methods=['GET', 'POST'])
def fix():
    form1 = ReusableForm(request.form)
    if request.method == 'POST':
        if form1.validate():
            bank = request.form.get('bank')
            save = request.form.get('save')
            institution = request.form.get('institution')
            rrn_column = (int(request.form['rrn_column'])) - 1
            pan_column = (int(request.form['pan_column'])) - 1
            amount_column = (int(request.form['amount_column'])) - 1
            file = request.files['file']
            if bank == None or institution == None or file.filename == '':
                flash('Error: All Fields are Required')
            else:
                global filename
                output_file_name = str(get_time() + "." + bank + "." + str(random.randint(100000, 999999)))
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                path = str(os.path.join(UPLOAD_FOLDER, str(filename)))
                if institution == 'ET-Switch':
                    main(save,institution,output_file_name,bank, path, rrn_column, pan_column, amount_column)
                    write_to_disk(filename, rrn_column, pan_column, amount_column)
                    flash('{}, [RRN Column = {}, PAN Column = {}, Amount Column = {}]'.format(filename, rrn_column,pan_column, amount_column))
                    return render_template('fix_result.html', rrn_column=rrn_column, pan_column=pan_column, amount_column=amount_column,bank=bank,output_file_name=output_file_name,filename=filename, institution=institution, save=save)
                else:
                    flash('Error: Institution not working!!')
        else:
            flash('Error: All Fields are Required')
    return render_template('fix.html', form1=form1)


@app.route("/", methods=['GET', 'POST'])
def index():
    return render_template('index.html')

@app.route('/download')
def download_help():
    try:
        path = "YCO_Help.pdf"
        return send_file(path, download_name = str("YCO_Help.pdf"), as_attachment=True)
    except:
        return render_template('error.html', error = 'Download Error!!')

@app.route('/log', methods=['GET', 'POST'])
def log():
    if request.method == 'POST':
        update_log()
        flash('Success: Up to date!!')
        return render_template(('log.html'))
    else:
        return render_template('log.html')

# @app.route('/callback', methods=['GET', 'POST'])
# def log():
#     if request.method == 'POST':
#         return 'OK'


@app.route("/logdesp", methods=["POST", "GET"])
def logdesp():
    try:
        mydb = mysql.connector.connect(host="localhost",port="3308", user="root", password="root", database="settle")
        mycursor = mydb.cursor(buffered=True)
        if request.method == 'POST':
            draw = request.form['draw']
            row = int(request.form['start'])
            rowperpage = int(request.form['length'])
            searchValue = request.form['search[value]']

            ## Total number of records without filtering
            mycursor.execute("select count(*) from settle_all")
            rsallcount = mycursor.fetchone()
            totalRecords = rsallcount[0]

            ## Total number of records with filtering
            likeString = "%" + searchValue + "%"
            mycursor.execute(
                "SELECT count(*) from settle_all WHERE RRN LIKE %s OR AQRRN LIKE %s OR PAN LIKE %s",
                (likeString, likeString, likeString))
            rsallcount = mycursor.fetchone()
            totalRecordwithFilter = rsallcount[0]

            ## Fetch records
            if searchValue == '':
                mycursor.execute("SELECT * FROM settle_all ORDER BY INSTITUTION asc limit %s, %s;", (row, rowperpage))
                settle_all_list = mycursor.fetchall()
            else:
                mycursor.execute(
                    "SELECT * FROM settle_all WHERE FEEDBACK LIKE %s OR STATUS_CLOSED LIKE %s OR BANK LIKE %s OR FILE_NAME LIKE %s OR STATUS LIKE %s OR RRN LIKE %s OR AQRRN LIKE %s OR PAN LIKE %s limit %s, %s;",
                    (likeString, likeString, likeString, likeString, likeString,likeString, likeString, likeString, row, rowperpage))
                settle_all_list = mycursor.fetchall()

            data = []
            for row in settle_all_list:
                data.append({
                    'FILE_NAME': row[1],
                    'INSTITUTION': row[2],
                    'BANK': row[3],
                    'RRN': row[4],
                    'PAN': row[5],
                    'AMOUNT': row[6],
                    'AQRRN': row[7],
                    'APPROVED_ON': get_date_format(row[8]),
                    'INCLUDED_IN_OUTGOING': get_date_format(row[9]),
                    'CLEARING_DATE': get_date_format(row[10]),
                    'FEEDBACK': row[11],
                    'ADDITIONAL_FEEDBACK': row[12],
                    'REQUESTED_DATE': get_date_format(row[13]),
                    'STATUS_CLOSED': get_date_format(row[14]),
                    'STATUS': row[15],
                })

            response = {
                'draw': draw,
                'iTotalRecords': totalRecords,
                'iTotalDisplayRecords': totalRecordwithFilter,
                'aaData': data,
            }
            return jsonify(response)
    except Exception as e:
        print(e)
    finally:
        mycursor.close()

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0',port='5000',threaded=True)
    # app.run(debug=True)