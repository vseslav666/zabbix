from pyzabbix import ZabbixAPI
import pandas
import datetime
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os



#функция певеродящая время работ в проценты от месяцы
def impact_time(start_time, end_time):
    start_time_sec = start_time[13:15]
    start_time_min = start_time[10:12]
    start_time_hour = start_time[8:10]
    start_time_day = start_time[5:7]

    end_time_sec = end_time[13:15]
    end_time_min = end_time[10:12]
    end_time_hour = end_time[8:10]
    end_time_day = end_time[5:7]

    start_time_sec = (int(start_time_day)*86400) + int(start_time_hour)*3600 + int(start_time_min)*60 + int(start_time_sec)
    end_time_sec = (int(end_time_day)*86400) + int(end_time_hour)*3600 + int(end_time_min)*60 + int(end_time_sec)
    a = (((end_time_sec - start_time_sec)*100)/43200)
    return a

#Функция проверяющая наличее файла
def create_impact_file():
    try:
        df1 = pandas.read_excel("impact_time.xls", index_col=0)
    except FileNotFoundError:
        df1 = pandas.DataFrame(columns = ['Имя устройства','Тип сбоя','Продолжительность сбоя', 'Время начала сбоя'])
        df1.to_excel("impact_time.xls")
#Функция создающая файл со времением сбоя
def write_impact(hostname, impact_type, time_impact, start_time):
    df1 = pandas.read_excel("impact_time.xls", index_col=0)
    if (len(df1)) == 0:
        df1.loc[0] = (hostname, impact_type, time_impact, start_time)
    elif (len(df1)) > 0:
            df1.loc[len(df1)+1] = (hostname, impact_type, time_impact, start_time)
    df1.to_excel("impact_time.xls")



def create_table_report(company):
    if company == 17:
        filename = ('ESB_SLA_Report ' + str(datetime.date.today()) + '.xlsx')
    elif company == 18:
        filename = ('SSK_SLA_Report ' + str(datetime.date.today()) + '.xlsx')
    z = ZabbixAPI('http://10.132.18.34/zabbix')
    z.login('telecom','fe35bf68917aca4779df6')
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    impact = pandas.read_excel("impact_time.xls", index_col=0)
    upper_row = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 2, 'left':2, 'bottom':1})
    upper_row_1 = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 2, 'bottom':1, 'left':1, 'right':1})
    upper_row_2 = workbook.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 2, 'right':2, 'bottom':1})
    under_row = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 1, 'left':2, 'bottom':2})
    under_row_1 = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 1, 'bottom':2, 'left':1, 'right':1})
    under_row_2 = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 1, 'right':2, 'bottom':2})
    left_column = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 1, 'bottom':1, 'left':2, 'right':1})
    center_column = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 1, 'bottom':1, 'left':1, 'right':1})
    right_column = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'top', 'top': 1, 'bottom':1, 'left':1, 'right':2})

    worksheet.write('B3', 'Отделение', upper_row)
    worksheet.write('C3', 'Отсутсвие электропитания, %', upper_row_1)
    worksheet.write('D3', 'Авария оператора связи, %', upper_row_1)
    worksheet.write('E3', 'Недоступность по вине Интер РАО-ИТ, %', upper_row_1)
    worksheet.write('F3', 'Доступность, %', upper_row_2)




    hosts = z.host.get(groupids=company, output=['hostid','name'])
    available = ''
    provider = ''
    blackout = ''
    i=2

    for host in hosts:
        hostname = host['name']
        i+=1
        items = z.item.get(hostids=(host['hostid']), output=['itemid','name', 'lastvalue'])
        for item in items:
            if item['name'] == 'Общее время доступности узла ({$REPORTDAYS} дней)':
                available = item['lastvalue']
                continue
            elif item['name'] == 'Узел недоступен по вине оператора связи ({$REPORTDAYS} дней)':
                provider = item['lastvalue']
                continue
            elif item['name'] == 'Узел недоступен по электропитанию ({$REPORTDAYS} дней)':
                blackout = item['lastvalue']
                continue
            else:
                continue
        for impact_str in impact.itertuples():
            impact_time = round(impact_str[3], 4)
            if impact_str[1] == hostname:
                if impact_str[2] == 'power':
                    if i != len(hosts)+2:
                        worksheet.write(i, 1, hostname, left_column)
                        worksheet.write(i, 2, float(blackout) - impact_time, center_column)
                        worksheet.write(i, 3, provider, center_column)
                        worksheet.write(i, 4, impact_time, center_column)
                        worksheet.write(i, 5, available, right_column)
                        break
                    else:
                        worksheet.write(i, 1, hostname, under_row)
                        worksheet.write(i, 2, float(blackout) - impact_time, under_row_1)
                        worksheet.write(i, 3, provider, under_row_1)
                        worksheet.write(i, 4, impact_time, under_row_1)
                        worksheet.write(i, 5, available, under_row_2)
                        break
                elif impact_str[2] == 'network':
                    if i != len(hosts)+2:
                        worksheet.write(i, 1, hostname, left_column)
                        worksheet.write(i, 2, blackout, center_column)
                        worksheet.write(i, 3, float(provider)-impact_time, center_column)
                        worksheet.write(i, 4, impact_time, center_column)
                        worksheet.write(i, 5, available, right_column)
                        break
                    else:
                        worksheet.write(i, 1, hostname, under_row)
                        worksheet.write(i, 2, blackout, under_row_1)
                        worksheet.write(i, 3, float(provider)- impact_time, under_row_1)
                        worksheet.write(i, 4, impact_time, under_row_1)
                        worksheet.write(i, 5, available, under_row_2)
                        break
                elif impact_str[2] == 'maintain':
                    if i != len(hosts)+2:
                        worksheet.write(i, 1, hostname, left_column)
                        worksheet.write(i, 2, blackout, center_column)
                        worksheet.write(i, 3, float(provider)-impact_time, center_column)
                        worksheet.write(i, 4, impact_time, center_column)
                        worksheet.write(i, 5, available, right_column)
                        break
                    else:
                        worksheet.write(i, 1, hostname, under_row)
                        worksheet.write(i, 2, blackout, under_row_1)
                        worksheet.write(i, 3, provider, under_row_1)
                        worksheet.write(i, 4, impact_time, under_row_1)
                        worksheet.write(i, 5, float(available)-impact_time, under_row_2)
                        break
            else:
                if i != len(hosts)+2:
                    worksheet.write(i, 1, hostname, left_column)
                    worksheet.write(i, 2, blackout, center_column)
                    worksheet.write(i, 3, provider, center_column)
                    worksheet.write(i, 4, '0', center_column)
                    worksheet.write(i, 5, available, right_column)

                else:
                    worksheet.write(i, 1, hostname, under_row)
                    worksheet.write(i, 2, blackout, under_row_1)
                    worksheet.write(i, 3, provider, under_row_1)
                    worksheet.write(i, 4, '0', under_row_1)
                    worksheet.write(i, 5, available, under_row_2)



    worksheet.set_column('B:B', 25)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 20)
    worksheet.set_column('F:F', 20)
    worksheet.set_landscape()
    worksheet.fit_to_pages(1,1)
    workbook.close()

def send_email(company):
    if company == 17:
        filepath = ('ESB_SLA_Report ' + str(datetime.date.today()) + '.xlsx')
        sender = "NOC@interrao.ru"
        server = 'msk1-smtp.interrao.ru'
        reciever = 'shtin_vg@interrao.ru'
        subject = 'Энергосбыт Волга. Отчет по SLA'
    elif company == 18:
        filepath = ('SSK_SLA_Report ' + str(datetime.date.today()) + '.xlsx')
        sender = "NOC@interrao.ru"
        server = 'msk1-smtp.interrao.ru'
        reciever = 'shtin_vg@interrao.ru'
        subject = 'Северная Сбытовая компания. Отчет по SLA'


# Compose attachment
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(filepath,'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(filepath))

# Compose message
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = reciever
    msg['Subject'] = subject
    msg.attach(part)

# Send mail

    smtp = smtplib.SMTP(server)
    print(filepath)
#   smtp.set_debuglevel(1)
    smtp.connect(server)
    smtp.sendmail(sender, reciever, msg.as_string())
    smtp.quit()


date = (str(datetime.date.today()))
print(date[-2:])
create_table_report(17)
send_email(17)
create_table_report(18)
send_email(18)

if date[-2:] == '01':
    os.rename ('impact_time.xls', 'impact_time.xlsx' +  date + '.bak')
