from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import selenium
import os
import time
import win32com.client
from datetime import datetime, timezone
import xml
import urllib.request
import requests
import xml.etree.ElementTree as ET


#all validation shit
old_url_uname = 'zachary.shaver'
old_url_pword = 'Cba32101!!'
new_url_uname = 'zachary.shaver'
new_url_pword = 'Cba32101!!'
old_sname = 'baefed'
new_sname = 'baefed'
xml_url = '/WBXService/XMLService'
#urls
old_webex_url = 'https://baefed.webex.com'
new_webex_url = 'https://baefed.webex.com'

#paths for winium automation
ptoneclk_path = "C:/Program Files (x86)/WebEx/Productivity Tools/ptoneclk.exe"
outlook_path = "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Microsoft Office 2016/Outlook 2016.lnk"

#username for testing
users_name = 'Shaver, Zachary T (US)'
search_string = 'Organizer:(' + users_name + ')'
ol_add = 'zachary.shaver'
meetings = {}

weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

def api_attend(m_key):
    xml_data = '''<?xml version="1.0" encoding="UTF-8"?>
<serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:serv="https://www.webex.com/schemas/2002/06/service">
    <header>
        <securityContext>
          <siteName>'''
    xml_data+=old_sname
    xml_data+='''</siteName>
          <webExID>'''
    xml_data+=old_url_uname
    xml_data+='''</webExID>
          <password>'''
    xml_data+=old_url_pword
    xml_data+='''</password>            
        </securityContext>
    </header>
    <body>
        <bodyContent xsi:type="java:com.webex.service.binding.meeting.GetMeeting">
            <meetingKey>'''
    xml_data+=m_key.replace(" ", "")
    xml_data+='''</meetingKey>
        </bodyContent>
    </body>
</serv:message>'''
    #print(xml_data)
    headers = {"Content-Type": "application/xml"}
    data = xml_data.encode('UTF-8')
    response = requests.post(old_webex_url+xml_url, data=xml_data, headers=headers)
    xml_response = response.content
    tree = ET.fromstring(xml_response)
    print(xml_response)
    print(tree.find('attendee'))




def start_winium():
    help(selenium)
    os.startfile(r"C:\Users\Zachary.shaver\Downloads\Winium.Desktop.driver\Winium.Desktop.driver.exe")
    time.sleep(10)

def end_winium():
    os.system("TASKKILL /F /IM Winium.Desktop.w_driver.exe")

def add_webex_on_ol():
    # outlook winium object
    o_driver = webdriver.Remote(
        command_executor='http://localhost:9999',
        desired_capabilities={
            "debugConnectToRunningApp": 'false',
            "app": outlook_path
        })
    o_driver.find_element_by_name('Mail').click()
    time.sleep(5)
    o_driver.find_element_by_name('Schedule Meeting').click()
    print('sch meeting hit')

    try:
        o_driver.find_element_by_name('OK').click()
    except selenium.common.exceptions.NoSuchElementException:
        print('settings already set')

    o_driver.find_element_by_name('To').send_keys('Drum, Dylan (US);Stewart, Sean (US)')
    o_driver.find_elements_by_name('Subject')[1].send_keys('TEST MEETING sent by winium')
    o_driver.find_element_by_name('Start date').send_keys('6/29/2018')
    o_driver.find_elements_by_name('Start time')[1].send_keys('13:00')
    o_driver.find_element_by_name('End date').send_keys('6/29/2018')
    o_driver.find_elements_by_name('End time')[1].send_keys('15:00')
    o_driver.find_element_by_name('Body').send_keys('test')
    o_driver.find_element_by_name('Recurrence...').click()
    o_driver.find_element_by_name('OK').click()
    o_driver.find_element_by_name('Send').click()







def get_meetings():
    # win32com outlook for getting email calender items
    outlook = win32com.client.Dispatch('Outlook.Application')
    namesp = outlook.GetNamespace('MAPI')
    msc = win32com.client.constants
    print(vars(win32com.client.constants))

    cal_folder = namesp.GetDefaultFolder(9).Items
    z = 0

    for meeting in cal_folder:


        #empty case
        if len(meeting.Recipients) == 0:
            continue

        #temp dict to store in the mas dict with key as the webex meeting number
        meeting_holder = {'Subject': '',
                          'Recipients': '',
                          'Body': '',
                          'Recurrence': '',
                          'Rec Pattern': [],
                          'Rec Pattern Start': '',
                          'Rec Pattern End': '',
                          'Start': '',
                          'End': ''}

        query_string = 'Reminder'

        r = ''
        alist = []

        #gets all the Recipients and stores them in a delemiter string and in a list
        for x in meeting.Recipients:
            r+= (x.Name + '; ')
            alist.append(x.Name)

        #checks if the meeting is organized by the user
        if users_name.lower() in alist[0].lower():
            #subject of the meeting
            if 'TEST' not in meeting.Subject:
                continue
            api_attend(access_code(meeting.Body))

            continue
            print_meeting(meeting, r)

            #find if it is recurring and what type of recurrence
            if meeting.IsRecurring:
                pat = meeting.GetRecurrencePattern()
                rec = pat.RecurrenceType

                if rec == msc.olRecursWeekly:                   #weekly
                    print('weekly')

                    #for weekly meetings there is a bitstring that get the pattern
                    #start and end date stored
                    meeting_holder['Recurrence'] = 'Weekly'
                    meeting_holder['Rec Pattern Start'] = pat.PatternStartDate
                    #print(pat.PatternStartDate)
                    #print(pat.PatternEndDate)

                    #pattern converted and stored in a list
                    meeting_holder['Rec Pattern End'] = pat.PatternEndDate
                    meeting_holder['Rec Pattern'] = dowm_convert(pat)

                elif rec == msc.olRecursDaily:                  #daily
                    print('daily')
                    meeting_holder['Rec Pattern Start'] = pat.PatternStartDate
                    meeting_holder['Rec Pattern End'] = pat.PatternEndDate
                    meeting_holder['Recurrence'] = 'Daily'
                elif rec == msc.olRecursMonthly:                #month
                    print('monthly')
                    meeting_holder['Rec Pattern End'] = pat.PatternEndDate
                    meeting_holder['Rec Pattern End'] = pat.PatternEndDate
                    meeting_holder['Rec Pattern'] = pat.DayOfMonth
                    meeting_holder['Recurrence'] = 'Monthly'
                elif rec == msc.olRecursMonthNth:               #monthNTH
                    print('monthnth')

                    meeting_holder['Rec Pattern'].append({'Interval': pat.Interval,         #Interval of months (1 every month)
                                                          'DayOfWeek': pat.DayOfWeekMask,   #Day of week bin
                                                          'Instance': pat.Instance})        #the (Instance)th (DayOfWeek) every (Interval) Month(s)
                    #print(pat.Interval)
                    #print(pat.DayOfWeekMask)
                    #print(pat.Instance)
                    meeting_holder['Recurrence'] = 'Monthly'
                elif rec == msc.olRecursYearly:                 #yearly
                    print('yearly')
                    meeting_holder['Rec Pattern'].append({'MonthOfYear': pat.MonthOfYear,
                                                          'DayOfMonth': pat.DayOfMonth})
                    meeting_holder['Rec Pattern End'] = pat.PatternEndDate
                    meeting_holder['Rec Pattern End'] = pat.PatternEndDate
                    meeting_holder['Recurrence'] = 'Yearly'          #yearNTH
                elif rec == msc.olRecursYearNth:
                    print('yearnth')
                    meeting_holder['Rec Pattern'].append({'Instance': pat.Instance,
                                                          'DayOfWeek':  pat.DayOfWeekMask,
                                                          'MonthOfYear': pat.MonthOfYear})
                    meeting_holder['Recurrence'] = 'Yearly'

                meeting_holder['Subject'] = meeting.Subject
                meeting_holder['Recipients'] = r
                meeting_holder['Body'] = meeting.Body
                meeting_holder['Start'] = meeting.Start
                meeting_holder['End'] = meeting.End
                # "Reminder Yes, To Drum, Dylan (US), Subject TEST MEETING IGNORE (month rec standard), Sent 12:19 PM, Size 15 KB, Flag Status Unf
                if meeting.ReminderSet:
                    query_string += ' Yes, '
                else:
                    query_string += ' No, '
                #query_string += (r[:-2] + ', Subject ' + meeting.Subject  + ', Sent '

                meetings.update({access_code(meeting.Body): meeting_holder})

                #meeting.Delete()
                z += 1
            else:
                print('non rec')
                if is_future(meeting.Start):
                    meeting_holder['Recurrence'] = 'x'
                    meeting_holder['Subject'] = meeting.Subject
                    meeting_holder['Recipients'] = r
                    meeting_holder['Body'] = meeting.Body
                    meeting_holder['Start'] = meeting.Start
                    meeting_holder['End'] = meeting.End
                    meetings.update({access_code(meeting.Body): meeting_holder})
                    z += 1
                else:
                    print('past')

    print(z)

#checks if it is a future meeting
def is_future(start):
    return datetime.now(timezone.utc) < start

#converts day of week mask to a readable thing since its in bits
def dowm_convert(pat):
    rec_bit = list(map(int, reversed("{0:b}".format(pat.DayOfWeekMask))))
    for x in range(len(rec_bit), 7):
        rec_bit.append(0)
    # print(rec_bit)
    return rec_bit

#prints basic info about the meeting
def print_meeting(meeting, r):
    print(meeting.End)
    print(meeting.Start)
    print(access_code(meeting.Body))
    print(r)
    print(meeting.Subject)

#locates and returns the access code of the body of a webex meeting
def access_code(body):
    if old_webex_url in body:
        return body.split('(access code):')[1][1:13]
    else:
        return ''

#i mean if you really need it its here
def parse_email(name, email):
    if '@' in name:
        return name
    elif '@' in email and '/' not in email:
        return email
    pmail = email.split('/')[4].split('=')[1]
    if '@' in pmail:
        if len(pmail.split('.')[-1]) == 6:
            return pmail[:-3]
        else:
            return pmail
    elif '.' not in pmail:
        return name.split(' ')[1] + '.' + name.split(',')[0] + '@baesystems.com'

    add = '@baesystems.com'
    efname = pmail.split('.')[0]
    elname = pmail.split('.')[1]
    lname = name.split(',')[0]
    fname = name.split(' ')[1]
    mname = ''

    #cases where the first name is not the same as the email one
    if efname != fname:
        if len(lname.split(' ')) > 5:
            lname = elname
    if len(name.split(' ')) > 3:
        mname = name.split(' ')[2]
    if len(pmail.split('.')) > 2 and '(' not in mname:
        return (fname+'.'+mname+'.'+lname+add).lower()
    else:
        return (fname+'.'+lname+add).lower()

#for swaping to the new url
def swap_url():
    w_driver = webdriver.Remote(
        command_executor='http://localhost:9999',
        desired_capabilities={
            "debugConnectToRunningApp": 'false',
            "app": ptoneclk_path
        })
    time.sleep(5)

    w_driver.find_element_by_name('Settings').click()
    w_driver.find_element_by_name('Sign Out').click()
    time.sleep(5)
    w_driver.find_element_by_name('siteurlview').send_keys('baefed.webex.com')
    w_driver.find_element_by_name('Next').click()
    time.sleep(5)

    print('ENTER USERNAME AND PASSWORD IN THE WEBEX PRODUCTIVITY TOOLS AND SELECT NEXT')
    input()

    w_driver.find_element_by_name('Sign In').click()

    w_driver.close()



get_meetings()


