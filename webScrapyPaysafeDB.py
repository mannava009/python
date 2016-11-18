import datetime
import BeautifulSoup
from robobrowser import RoboBrowser
import pandas as pd
import sqlalchemy as sql



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
    
# Step6:Parsing HTML Table 
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
    record = (ReportDate, TransDate, CaseNo, CardNo, Amount, CBReasonCodeDesc)
    appended_data.append(record)

#Data preparation to load into DB's    
columns=("input_date","transaction_date","case_number","cardno","amount","reason_description")
ChargeBackDetails = pd.DataFrame(appended_data)   
ChargeBackDetails.columns = columns
ChargeBackDetails['input_date']=pd.to_datetime(ChargeBackDetails['input_date'])
ChargeBackDetails['transaction_date']=pd.to_datetime(ChargeBackDetails['transaction_date'])
ChargeBackDetails['cc_first_6']=ChargeBackDetails['cardno'].str[:6]
ChargeBackDetails['cc_last_4']=ChargeBackDetails['cardno'].str[-4:]
ChargeBackDetails['amount']=ChargeBackDetails['amount'].str.replace('\(|\)|\$','')
del ChargeBackDetails['cardno']


#Step7:Loading into Excel
#writer = pd.ExcelWriter('output\chargeback.xlsx')
#ChargeBackDetails.to_excel(writer, 'sheet1', index=False)
#writer.save()

#Step7:Loading into MYSQL Databse
'''engine = sql.create_engine("mysql+mysqlconnector://root:1Base1t@Ibasesqldb:3306/cb360automation")
ChargeBackDetails.to_sql('chargeback_automation',engine,if_exists='append',index=False)

#Step8:Loading into MongoDB
client = MongoClient("mongodb://192.168.203.116:27017")
db=client.cb360automation
paysafe_python=db.paysafe_python
paysafe_python.insert_many(ChargeBackDetails.to_dict('records'))

endTotalTime = datetime.datetime.now()
totalTime = endTotalTime - startTime
print 'Time in MilliSeconds for parsing and Loading into a Database:%s' % ((totalTime.total_seconds() * 1000))'''
