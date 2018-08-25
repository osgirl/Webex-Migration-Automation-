# Webex-Migration-Automation-
Webex Migration Automation and other related scripts for my job at bae systems 

If anyone is looking for help on a way to move an entire company over from one webex platform to another please feel free to 
use these scripts, the code is uncommented so if you need any help understanding it let me know. 

The program that does the migration is called XML_TEST.py, names are not my strong suit. 

What this program does is use the win32com outlook client to parse through meeting objects in your outlook window. 
For each of these objects it determines if the meeting contains a string that is a link to join a webex meeting on your old platform.
From there it canceles and reschedules this meetings using xml requests to the webex api.
This progam would work best if ran on each users individual machines.
There are better ways to do this im sure but with the permissions i had durring my time at BAE Systems i was unable to find another way.
Contact me at zshaver1@uncc.edu if you need any help understanding this code, its free for you to use pass it off as your own if
you want i just would like to see this code save a company some money and time.
