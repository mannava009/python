import os
from robobrowser import RoboBrowser
import BeautifulSoup
import pandas as pd
import datetime
from flask import Flask

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
'__PREVIOUSPAGE':'z-bjoV_pmfCLfDctEXPThYp83YjxdASmWV1tJivrxSWJv1Oe9Dde3BLbBzEubqM4ZGZMe9DUCCOw1Kns8ewhPClqhNLVtUMHpDTg0oC-qd2_QD4puYZlzvyslRGb4oHIb44zemiJzg085VlGkDTRKAJWDW7JggQ81K7TBDMEsAQ1',
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
'ContentPlaceHolder1_window2_clientStat':'[[[[null,3,null,null,"430px","260px",1,null,null,1,null,3]],[[[[[null,"Adjustment Details",null]],[[[[[]],[],[]],[{},[]],null],[[[[]],[],[]],[{},[]],null],[[[[]],[],[]],[{},[]],null]],[]],[{},[]],null],[[[[null,null,null,null]],[],[]],[{},[]],null]],[]],[{},[]],"3,3,,,430px,260px,0"]',
'_ig_def_dp_cal_clientState':'[[null,[],null],[{},[]],"01,2016,11"]',
'ctl00$_IG_CSS_LINKS_':'~/App_Themes/Blue/Blue.css|../../ig_res/Default/ig_monthcalendar.css|../../ig_res/Default/ig_dialogwindow.css|../../ig_res/Default/ig_texteditor.css|../../ig_res/Default/ig_shared.css'}
endTotalTime = datetime.datetime.now()
totalTime = endTotalTime - startTime
print 'Time in MilliSeconds for excel download Data Preparation:%s' % ((totalTime.total_seconds() * 1000))
# Step5:POST request for ExcelDownload
startTime = datetime.datetime.now()
excelResponse = browser.open(excelUrl, method="post", data=data)
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
    CBReasonCodeDesc = col[8].text.replace("&nbsp;",'')
    record = (ReportDate, TransDate, CaseNo, AuthCode, TransType, CBType, CardNo, Amount, CBReasonCodeDesc)
    appended_data.append(record)

columns = ('Report Date', 'Trans Date', 'Case No', 'Auth Code', 'Trans Type', 'CB Type', 'Card No', 'Amount', 'CBReasonCodeDesc')
ChargeBackDetails = pd.DataFrame(appended_data)   
ChargeBackDetails.columns = columns
writer=pd.ExcelWriter('C:\Users\karthikm\Desktop\chargeback.xlsx')
ChargeBackDetails.to_excel(writer,'sheet1',index=False)
writer.save()
endTotalTime = datetime.datetime.now()
totalTime = endTotalTime - startTime
print 'Time in MilliSeconds for parsing and Loading into a Dataframe:%s' % ((totalTime.total_seconds() * 1000))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
