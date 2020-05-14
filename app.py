import mysql.connector
from flask import Flask
from flask import Flask, flash, redirect, render_template, request, session, abort
import os
import pandas as pd
import json
import fpdf
from zipfile import ZipFile

app = Flask(__name__)
app.secret_key = "abc"


@app.route('/')
def loginPage():
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor()
    cursor.execute("CREATE DATABASE IF NOT EXISTS Application")
    cursor.execute(
        "create table IF NOT EXISTS Application.USER_DETAILS (userid varchar(50) PRIMARY KEY,password varchar(20),role varchar(20),status varchar(20))")
    cursor.execute(
        "create table IF NOT EXISTS Application.Employee (employee_id varchar(50) PRIMARY KEY,employee_name varchar(20),department varchar(20),designation varchar(20),salary int,month varchar(20),address varchar(50),phoneno int,email_id varchar(50),reporting_manager varchar(50),invoice_no int,year int)")
    # cursor.execute("insert into Application.USER_DETAILS (userid,password,role,status) values('admin','admin','superadmin','active')")
    cursor.execute(
        "INSERT INTO Application.USER_DETAILS (userid,password,role,status) VALUES ('admin','admin','superadmin','active') ON DUPLICATE KEY UPDATE userid='admin',password='admin',role='superadmin',status='active'")
    db.commit()
    db.close()
    if 'logged_in' not in session:
        return render_template('/login.html')
    else:
        return render_template('/home.html')


@app.route('/login', methods=['POST'])
def handleLogin():
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor()
    cursor.execute("SELECT COUNT(*) FROM Application.USER_DETAILS WHERE userid=%s AND password=%s",
                   (str(request.form['username']), str(request.form['password'])))
    record = cursor.fetchone()
    db.close()
    if int(record[0]) > 0:
        session['logged_in'] = True
        return render_template('/home.html')
    else:
        return render_template('/login.html', errorMessage="Invalid User Id & Password")


@app.route('/logout')
def logout():
    session.clear()
    return render_template('login.html', errorMessage="Logged Out Successfully")


@app.route('/resetPassword')
def resetPassword():
    return render_template('login.html')


@app.route('/fileUploadPage')
def fileUploadPage():
    return render_template('fileupload.html')


@app.route('/getUserList')
def getUserList():
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT userid,password,role,status FROM Application.USER_DETAILS")
    records = cursor.fetchall()
    db.close()
    return render_template('subuser.html', userList=records)


@app.route('/addUpdateUser')
def addUpdateUser():
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor()
    cursor.execute(
        "INSERT INTO Application.USER_DETAILS (userid,password,role,status) VALUES ('admin','admin','superadmin','active') ON DUPLICATE KEY UPDATE userid='admin',password='admin',role='superadmin',status='active'")
    db.close()
    return render_template('subuser.html', userList=records)


@app.route('/deleteUser')
def deleteUser():
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor()
    cursor.execute("delete Application.USER_DETAILS where userid=%s", str(request.form['username']))
    db.close()


@app.route('/getFilteredUserData')
def getFilteredUserData():
    query = "select * from Application.USER_DETAILS"
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor()
    cursor.execute("where month=%s", str(request.form['username']))
    db.close()


@app.route('/handleUpload', methods=['POST'])
def handleUpload():
    try:
        f = request.files['file']
        data_xls = pd.read_excel(f)
        jsonData = data_xls.to_json(orient='records')
        jsonData = json.loads(jsonData)
        validateFile(jsonData)
        for data in jsonData:
            print(data['employee_id'])
            db = mysql.connector.connect(host="localhost", user="root", passwd="")
            cursor = db.cursor()
            cursor.execute(
                "INSERT INTO Application.Employee (employee_id,employee_name,department,designation,salary,month,address,phoneno,email_id,reporting_manager,invoice_no,year)  VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (data['employee_id'], data['employee_name'], data['department'], data['designation'],
                 int(data['salary']), (data['month']), data['address'], int(data['phoneno']), data['email_id'],
                 data['reporting_manager'], int(data['invoice_no']), int(data['year'])))
            db.commit()
            db.close()
            pdf = fpdf.FPDF(format='letter')  # pdf format
            pdf.add_page()  # create new page
            pdf.set_font("Arial", size=12)  # font and textsize
            pdf.cell(200, 10, txt="Invoice Report", ln=1, align="C")

            for key, value in data.items():
                pdf.cell(200, 10, txt=str(key) + " : " + str(value), ln=1, align="L")
                pdf.output(str(data['month']) + "_" + str(data['employee_name']) + ".pdf")
            with ZipFile('Invoice.zip', 'w') as zip:
                for data in jsonData:
                    zip.write(str(data['month']) + "_" + str(data['employee_name']) + ".pdf")
                    os.remove(str(data['month']) + "_" + str(data['employee_name']) + ".pdf")
            zip.close()

            # print(jsonData['jsonData'])
            return render_template('fileupload.html', message="File Uploaded Successfully & Data Saved")
    except mysql.connector.Error as err:
        return render_template('fileupload.html', errorMessage="Database Exception" + str(err))

    except Exception as e:
        print(e)
        return render_template('fileupload.html',
                               errorMessage="Invalid file, Upload valid file in xls or xlsx format only")


def validateFile(jsonData):
    if 'employee_id' in jsonData and 'employee_name' in jsonData and 'department' in jsonData and 'designation' in jsonData and 'salary' in jsonData and 'month' in jsonData and 'address' in jsonData and 'phoneno' in jsonData and 'email_id' in jsonData and 'reporting_manager' in jsonData and 'invoice_no' in jsonData and 'year' in jsonData:
        return True
    else:
        return render_template('fileupload.html', errorMessage="Invalid file, all fields are not present in excel file")


@app.route('/fetchEmployeeDetails')
def fetchEmployeeDetails():
    db = mysql.connector.connect(host="localhost", user="root", passwd="")
    cursor = db.cursor(dictionary=True)
    cursor.execute("SELECT * FROM Application.Employee")
    records = cursor.fetchall()
    db.close()
    return render_template('employeeList.html', employeeList=records)


if __name__ == "__main__":
    #session.clear()
    app.secret_key = os.urandom(12)
    app.run(debug=True, host='localhost', port=5000)

