import os
from robobrowser import RoboBrowser
import BeautifulSoup
import pandas as pd
import psycopg2
import datetime
from sqlalchemy import create_engine
from apscheduler.schedulers.blocking import BlockingScheduler

sched = BlockingScheduler()
@sched.scheduled_job('interval', seconds=10)
def paySafe():
    login_url = "https://merituspayment.com/merchants/frmLogin.aspx"
    # Step1 : GET Request for login page
    startTime = datetime.datetime.now()
    browser = RoboBrowser(history=True)
    browser.open(login_url)
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds for hitting the page:%s' % ((totalTime.total_seconds() * 1000))
    

    # Step2 : Setting Userid and password for login page
    startTime = datetime.datetime.now()
    LogIn = browser.get_form(id='form1')
    username = "42162"
    password = "16Comcastic4"
    LogIn["ctl00$ContentPlaceHolder1$txtLoginID"].value = username
    LogIn["ctl00$ContentPlaceHolder1$txtPassword"].value = password
    LogIn["ctl00$ContentPlaceHolder1$btnLogin"].value = "Log In"
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds for login Data Preparation:%s' % ((totalTime.total_seconds() * 1000))
    
    # Step3 : post request for login page
    startTime = datetime.datetime.now()
    browser.submit_form(LogIn)
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds to login into page:%s' % ((totalTime.total_seconds() * 1000))
    
    # Step4 : Setting Form parameters for Charge Back xls download
    startTime = datetime.datetime.now()
    excelUrl = "https://merituspayment.com/merchants/web/SecureReportForms/frmChargebackDetail.aspx?ct=0&dt=0&rd=0"
    data = {'__EVENTTARGET':'ctl00$ContentPlaceHolder1$btnExpExcelChargebackDay',
    '__EVENTARGUMENT':'',
    '__LASTFOCUS':'',
    '__VIEWSTATE':'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',
    '__VIEWSTATEGENERATOR':'78FC0077',
    '__VIEWSTATEENCRYPTED':'',
    '__PREVIOUSPAGE':'EycxxWOUmf7xPBWRdi6KTxWkK1kjnJlu-Gv6NwWq4MLwpQWGzFfCmCnufngiSgFJC-lANh4k7hKCFYCM4IGtnUazq0drd1WBRKrqsdqtoCkIMrcmaTz8mij1hA4f8U8NxfF20-tTt1eYypZMojDw5rivtB55Fe4aAYKcOGhN3Ls1',
    'ContentPlaceHolder1_wdcFromDate_clientState':'|0|012016-10-1-0-0-0-0||[[[[]],[],[]],[{},[]],"012016-10-1-0-0-0-0"]',
    'ContentPlaceHolder1_wdcToDate_clientState':'|0|012016-11-30-0-0-0-0||[[[[]],[],[]],[{},[]],"012016-11-30-0-0-0-0"]',
    'ctl00$ContentPlaceHolder1$ddlCBType':'-1',
    'ctl00$ContentPlaceHolder1$ddlTransType':'-1',
    'ctl00$ContentPlaceHolder1$ddlCardType':'-1',
    'ContentPlaceHolder1_wceTransAmount_clientState':'|0|01||[[[[]],[],[]],[{},[]],"01"]',
    'ContentPlaceHolder1_wceTransAmount':'',
    'ContentPlaceHolder1_wteAuthCode_clientState':'|0|01||[[[[]],[],[]],[{},[]],"01"]',
    'ContentPlaceHolder1_wteAuthCode':'',
    'ctl00$ContentPlaceHolder1$tbxLastFour':'',
    'ctl00$ContentPlaceHolder1$rdExport':'1',
    'ctl00$ContentPlaceHolder1$cboPageSize':'10',
    'ContentPlaceHolder1_window2_clientState':'[[[[null,3,null,null,"430px","260px",1,null,null,1,null,3]],[[[[[null,"Adjustment Details",null]],[[[[[]],[],[]],[{},[]],null],[[[[]],[],[]],[{},[]],null],[[[[]],[],[]],[{},[]],null]],[]],[{},[]],null],[[[[null,null,null,null]],[],[]],[{},[]],null]],[]],[{},[]],"3,3,,,430px,260px,0"]',
    '_ig_def_dp_cal_clientState':'[[null,[],null],[{},[]],"01,2016,11"]',
    'ctl00$_IG_CSS_LINKS_':'~/App_Themes/Blue/Blue.css|../../ig_res/Default/ig_monthcalendar.css|../../ig_res/Default/ig_dialogwindow.css|../../ig_res/Default/ig_texteditor.css|../../ig_res/Default/ig_shared.css'}
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds for excel download Data Preparation:%s' % ((totalTime.total_seconds() * 1000))
    # Step5:POST request for ExcelDownload
    startTime = datetime.datetime.now()
    browser.open(excelUrl, method="post", data=data)
    content = browser.response.content
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds for excel download:%s' % ((totalTime.total_seconds() * 1000))
    
    # Parsing HTML Table and Loading it into a excel
    startTime = datetime.datetime.now()
    soup = BeautifulSoup.BeautifulSoup(content)
    table = soup.find("table", border=1)
    appended_data = []
    for row in table.findAll('tr')[1:-1]:
        col = row.findAll('td')
        ReportDate = col[0].text
        TransDate = col[1].text
        CaseNo = col[2].text
        AuthCode = col[3].text
        TransType = col[4].text
        CBType = col[5].text
        CardNo = col[6].text
        Amount = col[7].text
        CBReasonCodeDesc = col[8].text.replace("&nbsp;", '')
        record = (ReportDate, TransDate, CaseNo, AuthCode, TransType, CBType, CardNo, Amount, CBReasonCodeDesc)
        appended_data.append(record)
    
    columns = ('Report Date', 'Trans Date', 'Case No', 'Auth Code', 'Trans Type', 'CB Type', 'Card No', 'Amount', 'CBReasonCodeDesc')
    ChargeBackDetails = pd.DataFrame(appended_data)   
    ChargeBackDetails.columns = columns
    #writer = pd.ExcelWriter('C:\Users\karthikm\Desktop\chargeback.xlsx')
    #ChargeBackDetails.to_excel(writer, 'sheet1', index=False)
    #writer.save()
    engine=create_engine("postgres://gsyorjbgdnmaex:pT_q3BALcPkVUuvpjYCOMYD1zv@ec2-23-21-238-76.compute-1.amazonaws.com:5432/dbmh4ld11aqgvi")
    ChargeBackDetails.to_sql("chargeback_details",engine,if_exists="replace")
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds for parsing and Loading into a Dataframe:%s' % ((totalTime.total_seconds() * 1000))
    return ChargeBackDetails
a=paySafe()
sched.start()    
