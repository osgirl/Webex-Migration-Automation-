import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application')
namesp = outlook.GetNamespace('MAPI')

cal_folder = namesp.GetDefaultFolder(9).Items

for meeting in cal_folder:
    if 'DELETE THIS' in meeting.Subject:
        meeting.Delete()