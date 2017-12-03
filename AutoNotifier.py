#%%tkinter timer def 지금 개발중인 가장 중요한 것
# debugging/logging
import coloredlogs, logging
import inspect

#logging.basicConfig(level=logging.DEBUG,format='%(asctime)s:%(levelname)s:%(message)s')
coloredlogs.install()

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(funcName)s:%(message)s')

log_file_handler = logging.FileHandler('AutoNotifier.log')
log_file_handler.setLevel(logging.INFO)
log_file_handler.setFormatter(formatter)

log_stream_handler = logging.StreamHandler()
log_stream_handler.setLevel(logging.INFO)
log_stream_handler.setFormatter(formatter)

logger.addHandler(log_file_handler)
logger.addHandler(log_stream_handler)

# update_clock()
import time

# create_widget()
import tkinter as tk
from tkinter import ttk
from tkinter import *

# file_manager(action)
import os
import shutil
import re

# login_diba()
from selenium import webdriver
from datetime import date

# excel_to_pd()/diva_collector()
import warnings; warnings.simplefilter("ignore")
import codecs
import pandas as pd
from bs4 import BeautifulSoup

# email(team)
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# filter_team(team)
import datetime as dt
from dateutil.relativedelta import relativedelta

download_path = os.path.join(os.getenv('USERPROFILE'), 'Downloads')
base_path = r'D:\DIVA'
full_path = ''
regex = r"IbiZa_Export_\d{8}.xls"

filter_word = ['Holder SB', '펌업 예정']

def create_widgets():
    logger.debug('function called')
    global lbl_CurrentTime
    global ent_ExportTime
    global cbo_TeamSelection
    global btnStartNew
    global cboT
    global methods

    # label
    lbl_CurrentTime = tk.Label(text='')
    lbl_CurrentTime.grid(columnspan=2)
    lbl_TeamSelection = tk.Label(text='Team Selection')
    lbl_TeamSelection.grid(column=0, row=1)
    lbl_ExportTime = tk.Label(text='Export Time')
    lbl_ExportTime.grid(column=0, row=2)

    # create combo
    cboT = ('gtac1', 'gtac2', 'na', 'nokia', 'japan', 'all')
    cbo_TeamSelection = ttk.Combobox(root, values=cboT, state='readonly', justify=CENTER)
    cbo_TeamSelection.grid(column=1, row=1, sticky=W)
    cbo_TeamSelection.current(5)

    # entry
    ent_ExportTime = tk.Entry(root)
    ent_ExportTime.grid(column=1, row=2, sticky=W)

    # create start button
    methods = ['Start with new Download', 'Start with Existing DIVA']
    for method in methods:
        btnStart = tk.Button(root, text = method, command = lambda m=method: login_diva(m))
        btnStart.grid(column=methods.index(method), row=3, sticky=N+S+E+W)

def update_clock():
    #logger.debug('function called')
    global now

    now = time.strftime("%H:%M:%S")
    #now = time.strftime("%M:%S")

    lbl_CurrentTime.configure(text=now)
    
    if now == ent_ExportTime.get() and time.strftime('%w') not in ('6', '0'):
        cbo_TeamSelection.set('all')
        login_diva(methods[0])
    if file_manager('find'):
        excel_to_pd()
        email(cbo_TeamSelection.get())
        file_manager('move')
        browser.quit()
        full_path = ''
    root.after(1000, update_clock)

def login_diva(method):

    logger.debug('function called')
    global browser

    usernameStr = 'ikmo'
    passwordStr = 'vertex22'
    browser_path = r'D:\Programming\Selenium\Chrome\chromedriver_2.30.exe'

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    #options.add_argument('--headless')

    browser = webdriver.Chrome(browser_path, chrome_options = options)
    browser.set_page_load_timeout(600)
    browser.get("http://bms.dasannetworks.com:5524/elisa_brand_new/index.php")

    # activate input (새로 만들어진 기능이로군)
    browser.find_element_by_xpath('//*[@id="guide"]').click()

    # fill in username and hit the next button
    username = browser.find_element_by_name('user_id')
    username.send_keys(usernameStr)
    password = browser.find_element_by_name('user_pw')
    password.send_keys(passwordStr)
    password = browser.find_element_by_name('user_pw')

    #login button
    browser.find_element_by_xpath("//*[@type='image']").click()

    ##export received button / testing purpose
    #browser.find_element_by_xpath("//*[@href='#Received']").click()
    time.sleep(5)

    browser.find_element_by_xpath("//*[@class='btn btn-warning dropdown-toggle  btn-sm']").click()
    time.sleep(2)

    browser.find_element_by_xpath("//*[@class='glyphicon glyphicon-globe']").click()
    time.sleep(2)

    if method == methods[0]:
        #export all
        browser.execute_script("document.for_export4.submit()")
    else:
        #copy existing DIVA to Download
        pass    #추가필요

def file_manager(action):
    logger.debug('function called')
    global full_path


    exist = False;

    for dirname, dirnames, filenames in os.walk(download_path):
            for filename in filenames:
                source = os.path.join(dirname, filename)
                if re.findall(regex, filename) and len(filename) == 25:
                    exist = True
                    full_path = source
                    if action == 'remove':
                        os.remove(source)
                    elif action == 'move':
                        shutil.move(source, os.path.join(base_path, filename))
                    elif action == 'find':
                        pass
    return exist

def excel_to_pd():
    logger.debug('function called')
    global sheet
    global links

    f=codecs.open(full_path, 'r', 'utf-8')
    diva_excel = f.read()

    soup = BeautifulSoup(diva_excel)
    links = soup.find_all('a')

    sheet = pd.read_html(diva_excel, header=1)
    sheet = sheet[0].dropna(axis=0, thresh=4) #axis 와 thresh 값이 뭔지 모르겠음

def diva_collector():    #DIVA link collector exapt BMS
    logger.debug('function called')
    global filtered_sheet

    diva_a=[]
    diva_href=[]

    for tag in links:
        link = tag.get('href',None)
        if 'show_bug_detail_diva' in link:
            diva_a.append(tag)
            diva_href.append(link)

    # Adding No. and Last Comment
    feedback = []
    no = []
    text = ''

    cnt = 0
    for i in filtered_sheet.index:
        cnt += 1
        filtered_sheet['Report ID'][i] = diva_a[i]
        browser.get(diva_href[i])
        html = browser.page_source
        bs = BeautifulSoup(html)

        div_warn = bs.find_all('div', {'class' : 'bs-callout bs-callout-warning'})
        div_info = bs.find_all('div', {'class' : 'bs-callout bs-callout-info'})

        if div_warn:
            text = stripWE(div_warn[0].text)
            fb = '<span style="color:blue">Customer feedback followed;<br></span>'
            fb += '<span style="color:black">'
            fb += text[:300]
            if len(text) > 300:
                fb += '<br> --omit(생략)--'
            fb += '</span>'
            feedback.append(fb)
        elif div_info:
            text = stripWE(div_info[0].text)
            fb = '<span style="color:green">'
            fb += 'Commented by TAC while sending email as below;<br>'
            fb += '메일 보내는 동안 TAC의 답변이 다음과 같이 나갔음;<br>'
            fb += '<p class="small" style="color:black">{}<br>'.format(text[:300])
            if len(text) > 300:
                fb += '--omit(생략)-- </p>'
            fb += '</span>'
            feedback.append(fb)
        else:
            fb = '<span style="color:red">'
            fb += 'No feedback detected. Please give the first message.<br>'
            fb += '(등록된 Feedback이 없음. 첫 Feedback을 남기세요)<br>'
            fb += '</span>'
            feedback.append(fb)

        no.append(cnt)

    filtered_sheet['Last Feedback'] = pd.Series(feedback).values
    filtered_sheet['No.'] = pd.Series(no).values
    if len(filtered_sheet) > 0:
        filtered_sheet = filtered_sheet[~filtered_sheet['Last Feedback'].str.contains(('|').join(filter_word))]

def email(team):
    logger.debug('function called')

    # me == my email address
    # you == recipient's email address
    me = "kyoungmo.in@dasanzhone.com"
    gtac1 = ['tac-oversea@dasannetworks.com']
    gtac2 = ['bosco.cho@dasanzhone.com', 'hoangminh.nguyen@dasanzhone.com', 'dung.le@dasanzhone.com', 'tuanminh.nguyen@dasanzhone.com', 'tho.luong@dasanzhone.com']
    na = ['lcg@dasannetworks.com', 'valloney@dasanzhone.com']
    nokia = ['ATCA_PA@dasannetworks.com']
    japan = ['dns_japant@dasannetworks.com']
    gtacHead = ['simon.park@g.dasanzhone.com']

    testAccount = ['ikmo@dasannetworks.com']

    recipients = gtac1 + gtac2 + na + nokia + japan
    cc = gtacHead

    #recipients = testAccount
    #cc = testAccount

    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "DIVA Statistics in 3 Months"
    msg['From'] = me
    msg['To'] = ', '.join(recipients)
    msg['Cc'] = ', '.join(cc)

    # Create the body of the message (a plain-text and an HTML version).
    hello = '<h3>Dear Teams.<br><br>Please refer to the DIVA follow up status in 3 Months.<br>(3개월간 DIVA follow up 현황 참고 하시기 바랍니다.)</h3>'
    bye = "<p>BR<br>KM</p>"
    old_width = pd.get_option('display.max_colwidth')

    #link줄어드는 것 방지 - 이거하지 않으면 링크가 제대로 보이지 않음
    pd.set_option('display.max_colwidth', -1) #동일 pd.options.display.max_colwidth = 100

    head = '<html><head><style>'
    #head += r'<link href="D:\Programming\Python Project\AutoNotifier\AutoNotifier\StyleSheet.css" type="text/css" rel="stylesheet">'
    head += codecs.open(r'D:\Programming\Python Project\AutoNotifier\StyleSheet.css', 'r', encoding='utf8').read()
    head += '</style></head>'

    body = head + '<body>' + hello

    if team == 'all':
        for i in range(0,5):
            filter_team(cboT[i])
            diva_collector()
            l = len(filtered_sheet.index)
            team_msg = '{}{}{} {}Team은 {} 개의 DIVA 답변이 필요합니다{}'.format('<h1>', cboT[i].upper(), '</h1>', '<h4>', l, '</h4>')
            body += team_msg
            if l > 0:
                body += filtered_sheet.to_html(na_rep='-', index=False, escape=False, col_space=75)
            else:
                team_msg = '{}{}{}'.format('<h2>', 'Well Done. Thank you for your effort:)', '</h2>')
                body += team_msg
            body += '<p>&nbsp;</p>'
    else:
        filter_team(team)
        diva_collector()
        l = len(filtered_sheet.index)
        team_msg = '{}{}{} {}Team은 {} 개의 DIVA 답변이 필요합니다{}'.format('<h1>', team.upper(), '</h1>', '<h4>', l, '</h4>')
        body += team_msg

        if l > 0:
                body += filtered_sheet.to_html(na_rep='-', index=False, escape=False, col_space=75)
        else:
            team_msg = '{}{}{}'.format('<h2>', 'Well Done. Thank you for your effort:)', '</h2>')
            body += team_msg
        body += '<p>&nbsp;</p>'

    body += bye
    body += '</body></html>'

    #link원복
    pd.set_option('display.max_colwidth', old_width)

    # Record the MIME types of both parts - text/plain and text/html.
    html_table = MIMEText(body, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    msg.attach(html_table)

    # Send the message via local SMTP server.
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.ehlo()
    s.starttls()
    s.login(me, 'rud8ahah')

    # sendmail function takes 3 arguments: sender's address, recipient's address
    # and message to send - here it is sent as one string.
    s.sendmail(me, recipients + cc, msg.as_string())
    s.quit()
    print(dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' email sent')

def filter_team(team):
    logger.debug('function called')
    global filtered_sheet

    ##%% Create variable with TRUE if last update is in 3 Months
    last_3_months = sheet['last update'] > (dt.datetime.now() + relativedelta(months=-3)).strftime("%Y-%m-%d %H:%M:%S")
    #last_3_months = sheet['last update'] > dt.datetime(2017,7,6,16,00,00).strftime("%Y-%m-%d %H:%M:%S")
    #last_3_months = sheet['last update'] > dt.datetime(2016,7,1,0,0,0).strftime("%Y-%m-%d %H:%M:%S")

    # List of filters

    list_dasan = ['@dasannetworks.com']
    list_tac2 = ['cheng@dasannetworks.com']
    list_test = ['Example', 'TEST', 'test']
    list_japan = ['ARTERIA','MITSUBISHI','UTStarcom']
    list_na = ['DASAN','FK','Evenflow']
    list_gtac2 = ['FIBRAIN','YTP']
    list_nokia = ['ATCA','ESB24&ESA40']
    list_vietnam = ['VFT','VNPT','VIETTEL','Viettel','VNTT','VIETNAM','ICTech']
    list_status = ['Developing','Issued','Received','Replied','Reported','Verifying']

    # Filter mask
    email_dasan= sheet['last touch'].str.contains('|'.join(list_dasan), na = False)
    email_tac2 = sheet['last touch'].str.contains('|'.join(list_tac2), na = False)
    contract_test = sheet['Report ID'].str.contains('|'.join(list_test), na = False)
    contract_japan = sheet['Report ID'].str.contains('|'.join(list_japan), na = False)
    contract_na = sheet['Report ID'].str.contains('|'.join(list_na), na = False)
    contract_gtac2 = sheet['Report ID'].str.contains('|'.join(list_gtac2), na = False)
    contract_nokia = sheet['Report ID'].str.contains('|'.join(list_nokia), na = False)
    contract_vietnam = sheet['Report ID'].str.contains('|'.join(list_vietnam), na = False)
    diva_status = sheet['DIVA Status'].str.match('|'.join(list_status), na = False)

    # Filter by team
    if team == 'gtac1':     #gtac1
        filtered_sheet = sheet[
            ~(contract_japan | contract_na | contract_gtac2 | contract_nokia | contract_vietnam | contract_test) &
            last_3_months &
            diva_status &
            (~email_dasan | email_tac2)
            ]

    if team == 'gtac2':     #gtac2
        filtered_sheet = sheet[
            contract_gtac2 &
            last_3_months &
            diva_status &
            (~email_dasan | email_tac2)
            ]

    if team == 'na':     #na
        filtered_sheet = sheet[
            contract_na &
            last_3_months &
            diva_status &
            (~email_dasan | email_tac2)
            ]

    if team == 'nokia':     #nokia
        filtered_sheet = sheet[
            contract_nokia &
            last_3_months &
            diva_status &
            (~email_dasan | email_tac2)
            ]

    if team == 'japan':     #japan
        filtered_sheet = sheet[
            contract_japan &
            last_3_months &
            diva_status &
            (~email_dasan | email_tac2)
            ]

    # Removing unnessasary columns
    cols_of_interest = ['No.', 'Report ID', 'Reg date', 'last update', 'Reporter', 'Receptionist', 'Title']
    filtered_sheet = filtered_sheet[cols_of_interest]

    #temporarily remove
    #del(filtered_sheet['Dealer'])
    #del(filtered_sheet['last touch'])

    print(dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + ' filtered {}'.format(team))

def stripWE(word):
    w = word.strip().replace('\n',' ').replace('\t','').replace('\r','').replace(u'\xa0', u' ')
    return re.sub('From:.*Subject:','', w)

root = tk.Tk()
file_manager('remove')   #제거필요
create_widgets()
update_clock()
root.mainloop()