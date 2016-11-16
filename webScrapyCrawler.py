import datetime
import BeautifulSoup
from robobrowser import RoboBrowser
import pandas as pd
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
    LogIn = browser.get_form()
    username = "31495"
    password = "1326C0mcastic"
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
            'ContentPlaceHolder1_wdcToDate_clientState':'|0|012016-10-31-0-0-0-0||[[[[]],[],[]],[{},[]],"012016-10-31-0-0-0-0"]',
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
            'ctl00$_IG_CSS_LINKS_':'~/App_Themes/Blue/Blue.css|../../ig_res/Default/ig_monthcalendar.css|../../ig_res/Default/ig_dialogwindow.css|../../ig_res/Default/ig_texteditor.css|../../ig_res/Default/ig_shared.css'
    }

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
    engine=create_engine("postgres://gsyorjbgdnmaex:pT_q3BALcPkVUuvpjYCOMYD1zv@ec2-23-21-238-76.compute-1.amazonaws.com:5432/dbmh4ld11aqgvi")
    ChargeBackDetails.to_sql("chargeback_details",engine,if_exists="replace")
    endTotalTime = datetime.datetime.now()
    totalTime = endTotalTime - startTime
    print 'Time in MilliSeconds for parsing and Loading into a Dataframe:%s' % ((totalTime.total_seconds() * 1000))
    return ChargeBackDetails

paySafe()

sched.start()    
