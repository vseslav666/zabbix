import pandas
from flask import Flask, render_template, request
from func_last import create_impact_file, impact_time, write_impact, create_table_report, send_email
import pandas


app=Flask(__name__)

@app.route('/')
def home():
    return render_template("index.html")

@app.route('/success', methods=['POST'])
def success():
    if request.method=='POST':
        startTime = request.form["work_start_date"] + request.form["work_start_time"]
        endTime = request.form["work_end_date"] + request.form["work_end_time"]
        hostname = request.form["device_name"]
        impactTime = impact_time(startTime, endTime)
        impact_type = request.form["impact_type"]
        create_impact_file()
        write_impact(hostname, impact_type, impactTime, startTime)
        print(startTime, endTime)
        return render_template("success.html")

    return render_template("success.html")


@app.route('/report')
def report():
    return render_template("report.html")

@app.route('/report_done', methods=['POST'])
def report_done():
    if request.method == 'POST':
        company = request.form["company"]
        if company == 'SSK':
            company = 18
        elif company == 'ESB':
            company = 17


    create_table_report(company)
    send_email(company)
    return render_template("report_done.html")


if __name__=="__main__":
    app.run(host = '0.0.0.0')

