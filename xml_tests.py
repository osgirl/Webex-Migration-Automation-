import requests
import random
import string
import xml.etree.ElementTree as ET
import win32com.client

#url xml requests are sent to on the baefed site
xml_url = 'https://baefed.webex.com/WBXService/XMLService'
webex_ID = 'zachary.shaver'
site_name = 'baefed'

#for storing meetings infomation used in parse tree
mas = {}
old_url = 'baefed.webex.com'
minfo = []

#takes in the meetings id, sends an XML request using it, Returns parsed data
def get_meeting(id):

    #formats request
    xml_request = '''<?xml version="1.0" encoding="UTF-8"?>
    <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlns:serv="https://www.webex.com/schemas/2002/06/service">
        <header>
            <securityContext>
              <siteName>baefed</siteName>
              <webExID>'''
    xml_request+=old_umane
    xml_request+='''</webExID>
              <password>'''
    xml_request+=old_passs
    xml_request+='''</password>            
            </securityContext>
        </header>
        <body>
            <bodyContent xsi:type="java:com.webex.service.binding.meeting.GetMeeting">
                <meetingKey>'''
    xml_request+=id
    xml_request+='''</meetingKey>
            </bodyContent>
        </body>
    </serv:message>'''

    #encode and send response
    headers = {"Content-Type": "application/xml"}
    response = requests.post(xml_url, data=xml_request, headers=headers)

    #decode and convert to xml for parsing
    response.raw.decode_content = True
    res = response.content

    #validate the email looking for success in the response
    if 'SUCCESS' in str(response.content):
        print('GRAB ID SUCCESS FOR', id)
    else:
        print('FAILED TO GRAB BASED OFF OF ID')

    #goes to the parse tree function passing the response as an xml tree object, id of the meeting, and the unformatted res
    parse_tree(ET.fromstring(res), id, str(response.content))

#formats data
def parse_tree(tree, id, xmlstr):

    #sets up subdicts for mas dict
    attendies = {}
    schedule = {}
    meetinfo = {}
    meetingkey = ''

    #gets id from response
    for x in tree.iter('{http://www.webex.com/schemas/2002/06/service/meeting}meetingkey'):
        meetingkey = x.text

    #checks for match
    if meetingkey == id:
        print('GOOD MATCH WITH', id)
    else:
        print('ID MATCH FAILED', id, meetingkey)

    #gets all the scheduling information
    for x in tree.iter("{http://www.webex.com/schemas/2002/06/service/meeting}schedule"):
        sch_holder = {}
        for z in x:
            sch_holder.update({z.tag.split("}")[1]: z.text})
        schedule.update({sch_holder['startDate']: sch_holder})

    #gets all the attendiee infomrmation
    for x in tree.iter('{http://www.webex.com/schemas/2002/06/service/attendee}person'):
        att_holder = {}
        for z in x:
            att_holder.update({z.tag.split("}")[1]: z.text})
        attendies.update({att_holder['name']: att_holder})

    #gets all the info about the meeting
    for x in tree.iter('{http://www.webex.com/schemas/2002/06/service/meeting}metaData'):
        info_holder = {}
        for z in x:
            info_holder.update({z.tag.split("}")[1]: z.text})
            #print(z.tag.split("}")[1], z.text)
        meetinfo.update({info_holder['confName']: info_holder})

    #storing recurr pattern
    recurr = ''

    #checks the raw string for a repeat xml member
    if '<meet:repeat>' in xmlstr:
        recurr = parse_rec(xmlstr.split('<meet:repeat>')[1].split('</meet:repeat>')[0])
        print('RECURRENCE PATTERN NOTED')

    #stores in mas dict
    mas.update({'subject': list(meetinfo.keys())[0],
              'startdate': list(schedule.keys())[0],
              'attendees': attendies,
              'meetinfo': meetinfo,
              'schedule': schedule,
              'repeat': recurr})




#makes meeting off gathered infomation
def sch_meet():

    #generates random meeting pass
    meet_pass = ''.join(random.choice(string.ascii_uppercase + string.digits) for _ in range(8))

    #generates a call in pass
    callinpass = str(random.randint(10000,99999))

    #forms xml request based off data
    xml_data = '''<?xml version="1.0" encoding="UTF-8"?>
    <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlns:serv="https://www.webex.com/schemas/2002/06/service">
        <header>
            <securityContext>
              <siteName>baefed</siteName>
              <webExID>'''
    xml_data+=new_uname                                                                 #username
    xml_data+='''</webExID>
              <password>'''
    xml_data+=new_pass                                                                  #password
    xml_data+='''</password>            
            </securityContext>
        </header>
        <body>       
            <bodyContent
                xsi:type="java:com.webex.service.binding.meeting.CreateMeeting">
                <accessControl>
                    <meetingPassword>'''
    xml_data+= meet_pass                                                                #meeting pass
    xml_data+='''</meetingPassword>
                </accessControl>
                <metaData>
                    <confName>'''
    xml_data+=mas['subject']                                                         #meeting subject
    xml_data+='''</confName>
                    <meetingType>'''
    xml_data+=mas['meetinfo'][mas['subject']]['meetingType']                          #meeting type (12200 is normal meet)
    xml_data+='''</meetingType>
                </metaData>
                <participants>
                    <attendees>'''
    for x in mas['attendees'].keys():                                                #adds all attediees
        xml_data+='<attendee><person><name>'
        xml_data+=(mas['attendees'][x]['name']+'</name><firstName>')                 #name
        xml_data+=(mas['attendees'][x]['firstName']+'</firstName><lastName>')        #first name
        xml_data+=(mas['attendees'][x]['lastName']+'</lastName><email>')             #last name
        xml_data+=(mas['attendees'][x]['email']+'</email></person></attendee>')      #email
    xml_data+='''</attendees>
                </participants>'''
    if mas['repeat']:
        xml_data += ('<repeat>' + mas['repeat'] + '</repeat>')
    xml_data+='''<enableOptions>
                    <chat>true</chat>
                    <poll>true</poll>
                    <audioVideo>true</audioVideo>
                </enableOptions>
                <schedule>
                    <startDate>'''
    xml_data+=mas['startdate']                                                              #start date
    xml_data+='''</startDate>
                    <openTime>'''
    xml_data+=mas['schedule'][mas['startdate']]['openTime']                               #open time
    xml_data+='''</openTime>
                    <joinTeleconfBeforeHost>'''
    xml_data+=mas['schedule'][mas['startdate']]['joinTeleconfBeforeHost']                 #join before host bool
    xml_data+='''</joinTeleconfBeforeHost>
                    <duration>'''
    xml_data+=mas['schedule'][mas['startdate']]['duration']                               #durration of meeting
    xml_data+='''</duration>
                    <timeZoneID>'''
    xml_data+=mas['schedule'][mas['startdate']]['timeZoneID']                             #timezone
    xml_data+='''</timeZoneID>
                </schedule>
                <telephony>
                    <telephonySupport>CALLIN</telephonySupport>
                    <extTelephonyDescription>
                        Call 1-844-800-2712, Passcode '''
    xml_data+=callinpass                                                                 #meeting call in pass
    xml_data+='''</extTelephonyDescription>
                </telephony>
            </bodyContent>
        </body>
    </serv:message>'''

    #sends request
    headers = {"Content-Type": "application/xml"}
    data = xml_data.encode('UTF-8')
    response = requests.post(xml_url, data=xml_data, headers=headers)

    #validates that the meeting was created
    if 'SUCCESS' in str(response.content):
        print('SCHEDULE SUCCESS FOR', mas['subject'])
    else:
        print('FAILED TO SCHEDULE MEETING')

    #adds meeting info for the email thats updated later
    minfo.append('1-844-800-2712')
    minfo.append(meet_pass)
    response.raw.decode_content = True
    return response.content

#parses the emails body to get the id for that meeting
def access_code(body):
    if old_url in body:
        return body.split('(access code):')[1][1:13]
    else:
        return ''

#
def parse_response(response):

    tree = ET.fromstring(response)
    m_key = list(tree.iter('{http://www.webex.com/schemas/2002/06/service/meeting}meetingkey'))[0].text
    get_url_xml = '''<?xml version="1.0" encoding="UTF-8"?>
    <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
        xmlns:serv="https://www.webex.com/schemas/2002/06/service">
        <header>
            <securityContext>
              <siteName>baefed</siteName>
              <webExID>'''
    get_url_xml+=new_uname
    get_url_xml+='''</webExID>
              <password>'''
    get_url_xml+=new_pass
    get_url_xml+='''</password>            
            </securityContext>
        </header>
        <body>
           <bodyContent
                xsi:type="java:com.webex.service.binding.meeting.GetjoinurlMeeting">
                <sessionKey>'''
    get_url_xml+=m_key
    get_url_xml+='''</sessionKey>
            </bodyContent>
        </body>
    </serv:message>'''

    #print(new_uname,new_pass,m_key)
    #print(get_url_xml)

    headers = {"Content-Type": "application/xml"}
    re = requests.post(xml_url, data=get_url_xml, headers=headers)

    #print(xml.dom.minidom.parseString(re.content).toprettyxml())
    if 'SUCCESS' in str(re.content):
        print('SUCCESS FOR THE URL GRAB')
    else:
        print('FAILED URL GRAB', re.content)
    re.raw.decode_content = True

    tree = ET.fromstring(re.content)
    return list(tree.iter('{http://www.webex.com/schemas/2002/06/service/meeting}joinMeetingURL'))[0].text






def parse_rec(rec):
    retrec = ''
    for x in rec.split('<meet:'):
        if '>0<' not in x and '/>' not in x:
            retrec+=(x+'<')
    return (retrec.replace('meet:', ''))[:-1]

#sends a xml with the credentials tp the url passed in sticks in the loop if they get it wrong
def validate_url(url):
    response = ''

    #keeps them in until they get it right
    while 'SUCCESS' not in response:
        print(('Enter Your Credentials for the ' + url + '.webex.com Site'))
        uname = input('Enter Username: ')
        passs = input('Enter Password: ')
        auth_url_xml = '''<?xml version="1.0" encoding="UTF-8"?>
            <serv:message xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xmlns:serv="https://www.webex.com/schemas/2002/06/service">
                <header>
                    <securityContext>
                      <siteName>'''
        auth_url_xml+=url
        auth_url_xml+='''</siteName>
                      <webExID>'''
        auth_url_xml += uname
        auth_url_xml += '''</webExID>
                      <password>'''
        auth_url_xml += passs
        auth_url_xml += '''</password>            
                    </securityContext>
                </header>
    <body>
        <bodyContent xsi:type="java:com.webex.service.binding.user.AuthenticateUser">
        </bodyContent>
    </body>
</serv:message>'''

        #sends formatted xml
        headers = {"Content-Type": "application/xml"}
        re = requests.post(('https://'+url+'.webex.com/WBXService/XMLService'), data=auth_url_xml, headers=headers)

        #checks for success response and breaks out if so
        response = str(re.content)
        if 'SUCCESS' in response:
            print('***SUCCESSFUL PASSWORD VALIDATION***')
            return [uname, passs]
        print('***INCORRECT USERNAME OR PASSWORD PLEASE TRY AGAIN***')





#win32com objects
outlook = win32com.client.Dispatch('Outlook.Application')
namesp = outlook.GetNamespace('MAPI')
cal_folder = namesp.GetDefaultFolder(9).Items

#urls for the new email always the same so i stored them
callin_url = 'https://baefed.webex.com/baefed/globalcallin.php?serviceType=MC&ED=6687087&tollFree=1'
callinres_url = 'https://www.webex.com/pdf/tollfree_restrictions.pdf'
cantjoin = 'https://help.webex.com/docs/DOC-5412'

#gets the username and sends a welcome message
user = outlook.Session.CurrentUser.Name
print('Welcome, ', user)

#validation for both urls
old_umane, old_passs = validate_url('baefed')
new_uname, new_pass = validate_url('baefed')

#itterates trough meetings folder
for meeting in cal_folder:

    #the old url is in the body and they created the meeting
    if old_url in meeting.Body and 'test' in meeting.Subject.lower() and meeting.Organizer == user:

        #gets the access code of the meeting
        m_uniqid = access_code(meeting.Body)
        if not m_uniqid:
            continue

        #parse it because the xml request wants it will no spaces and get the information about the meeting
        get_meeting(m_uniqid.replace(" ", ""))

        #schedules meeting based off of the info from the get meeting
        response = sch_meet()

        #parses the response from that email so that it can get the join url for the meeting
        join_url = parse_response(response)

        #this is where the body of the email is updated
        body = meeting.Body

        #replaces strings by cutting the string by delims and then replaces it with the new info
        body = body.replace(body.split('Join WebEx meeting <')[1].split('>')[0], join_url)
        body = body.replace(body.split('Meeting number (access code): ')[1].split('M')[0], (m_uniqid+'\n'))
        body = body.replace(body.split('Meeting password: ')[1].split(' ')[0], (minfo[1] + '\n'))
        body = body.replace(body.split('Join by phone')[0].split(' U')[1], (minfo[0] + '\n'))
        body = body.replace(body.split('Global call-in numbers <')[1].split('>')[0], (callin_url))
        body = body.replace(body.split("Can't join the meeting? <")[1].split('>')[0], (callinres_url))

        #it normally only updates when time chenge this makes it update to any change
        meeting.ForceUpdateToAllAttendees = True

        #sets the formatted new body
        meeting.Body = body

        #saves the email object
        meeting.Save()

        #sends everyone update
        meeting.Send()

        #resets that one list from before bc of multiple meetings
        minfo = []

#done
print('***MIGRATION PROCESS COMPLETED PLEASE CLOSE THE PROGRAM***')