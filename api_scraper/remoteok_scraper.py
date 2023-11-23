import requests
import xlwt 
from xlwt import Workbook
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE,formatdate

#https://myaccount.google.com/lesssecureapps
API_URL='https://remoteok.com/api'
USER_AGENT='https://explore.whatismybrowser.com/useragents/parse/?analyse-my-user-agent=yes' #change my user agent to avoid any kind of problem of security
REQUEST_HEADER={
    'User-Agent': USER_AGENT,
    'Accept-language': 'en-US,en;q=0.5',
}



def get_job_posting():
    res=requests.get(url=API_URL,headers=REQUEST_HEADER)
    return res.json()

def output_jobs_to_xls(data):
    wb=Workbook() #instance of workbook
    sheet=wb.add_sheet('Jobs')
    xl_headers=list(data[0].keys())
    for i in range(len(xl_headers)):
        sheet.write(0,i,xl_headers[i]) #write (row index , column index,data)   
    for i in range(0,len(data)):
        job=data[i]
        values=list(job.values())
        for x in range (0,len(values)):
            sheet.write(i+1,x,values[x])
    wb.save('./api_scraper/remote_jobs.xls')

def send_email(send_form,send_to,subject,text,files=None):
    assert isinstance(send_to,list)
    msg=MIMEMultipart()
    msg['From']=send_form
    msg['to']=COMMASPACE.join(send_to)
    msg['date']=formatdate(localtime=True)
    msg['Subject']=subject

    msg.attach(MIMEText(text))
    for f in files or []:
        with open(f,"rb") as file:
            part=MIMEApplication(file.read(),Name=basename(f))
        part['Content-Disposition'] = f'attachment; filename="{basename(f)}"'
        msg.attach(part)

    smtp = smtplib.SMTP('smtp.gmail.com')
    smtp.starttls()
    smtp.login(send_form,'your_password')
    smtp.sendmail(send_form,send_to,msg.as_string())
    smtp.close()


if __name__=="__main__":
    json=get_job_posting()[1:] ##index[0]dummy informations we dont need them 
    output_jobs_to_xls(json)
    send_email('nadi.hanane01@gmail.com',['Contact@Meow.com'],'jobs posting','hey',files=['./api_scraper/remote_jobs.xls'])
