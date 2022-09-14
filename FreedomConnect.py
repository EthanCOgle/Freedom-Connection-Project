import time
#import pyautogui
import subprocess
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Imports

LOCATION = "512"
Device = "Node 3"
NumOfRows = 48
NumOfPings = "600"
NumOfDays = 1

#Global variables

pyautogui.FAILSAFE = False

#Turning off a failsafe for pyautogui where it will quit if mouse goes toward edge of screen

def sendMessage(packetLossPercent, reciever="regulatory"): #Sends a message to the team and mentors
    email = "################@outlook.com"
    pas = "************"
   
    #Freedom Connect outlook account since google does not allow thrid parties to log in
   
    if reciever == "******":
        sms_gateway = '###########@vtext.com'
    if reciever == "******":
        sms_gateway = '###########@mms.att.net'
    if reciever == "regulatory":
        sms_gateway = '###########@txt.att.net' # Change number for next generation
       
    #Sets the sms_gateway which allows an email to send a text message

    smtp = "smtp-mail.outlook.com"
    port = 587
   
    #Specific smtp and port number for outlook

    server = smtplib.SMTP(smtp,port)

    server.starttls()

    server.login(email,pas)

    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = sms_gateway
   
    #                                Determines message that is being sent out
    if packetLossPercent == 100.0: # If the internet goes down  
        body = 'The internet was down!'
        msg['Subject'] = 'Danger Level: 10 | Location: ' + LOCATION
    if packetLossPercent >= 0.0 and packetLossPercent < 1.0: # If the packet loss is at a normal level
        body = 'The internet is currently up and running and doing fine.'
        msg['Subject'] = 'Location: ' + LOCATION + ' | Packet Loss: ' + str(packetLossPercent)
    if packetLossPercent >= 1.0 and packetLossPercent < 2.5: # If the packet loss is slowing down
        body = 'The internet is currently up and running but is starting to slow down.'
        msg['Subject'] = 'Location: ' + LOCATION + ' | Packet Loss: ' + str(packetLossPercent)
    if packetLossPercent >= 2.5: # If the packet loss is abnormally high
        body = 'The internet is currently having trouble and slow.'
        msg['Subject'] = 'Location: ' + LOCATION + ' | Packet Loss: ' + str(packetLossPercent)
 
    msg.attach(MIMEText(body, 'plain'))

    sms = msg.as_string()
 
    server.sendmail(email, sms_gateway,sms)

    server.quit()
       
    print('Mail Sent')
   
   
def openUp(): # Opens up the spread sheet where we record all of the data. directly acceses the shell through subprocess
    subprocess.Popen(["libreoffice --calc /home/pi/Documents/FreedomConnectData.ods"], stdout=subprocess.DEVNULL, shell=True)
   
def pinging(): # Pings the internet (google.com) and returns the packet loss, minimum ping, average ping, and maximum ping
   
    pyautogui.hotkey('ctrl', 's')
   
    #to save each time while pinging
   
    try:
        Pinging = subprocess.check_output("ping -c "+NumOfPings+" google.com", shell=True) # (1 ping/sec)
        ping = str(Pinging)
       
        # Once again using a shell command with subprocess to ping the internet
       
        beforePercent = (ping.index("ved, ")) + 5
        afterPercent = (ping.index("% packet loss"))
        packetLoss = (ping[beforePercent:afterPercent])
       
        pings = ping[ping.index("mdev = ")+7:]
       
        minPing = pings[:pings.index("/")]
       
        pings = pings[len(minPing)+1:]
       
        avgPing = pings[:pings.index("/")]
       
        pings = pings[len(avgPing)+1:]
       
        maxPing = pings[:pings.index("/")]
       
        data = packetLoss + "+" + minPing + "-" + avgPing + ":" + maxPing
       
    except:
        data = "ERROR+ERROR-ERROR:ERROR"
        # If the internet is compeletly down, the subprocess shell command will return an error,
        # this try/except block will keep the program runnning and let the team know by entering in
        # 'ERROR' for that particular time slot/ row
       
    return data

def settingUp(): # This sets up the spread sheet by labeling the columns with their respective data types
   
    pyautogui.typewrite("Date/Time", interval=0)
   
    pyautogui.typewrite(['tab'], interval=0)
    pyautogui.typewrite("Packet Loss", interval=0)
   
    pyautogui.typewrite(['tab'], interval=0)
    pyautogui.typewrite("Min Ping(ms)", interval=0)
   
    pyautogui.typewrite(['tab'], interval=0)
    pyautogui.typewrite("Average Ping(ms)", interval=0)
   
    pyautogui.typewrite(['tab'], interval=0)
    pyautogui.typewrite("Max Ping(ms)", interval=0)

    pyautogui.typewrite(['tab'], interval=0)
    pyautogui.typewrite("Location", interval=0)
   
    pyautogui.typewrite(['tab'], interval=0)
    pyautogui.typewrite("Device", interval=0)
    for i in range(6):
        pyautogui.hotkey('shift','tab')
   
    pyautogui.typewrite(['enter'], interval=0)
   
    #Uses pyautogui which is a python package that gives control over the keyboard and mouse along with some other functions
    #To install, type 'python3 -m pip install pyautogui' into a shell
   
def eachRow(): # This enters in all of the data collected by the pinging() function along with the date/time, location,
               # and device number into the spread sheet.
    for i in range(6):
        pyautogui.typewrite(['tab'], interval=0)
    data = pinging() # calls pinging() function that returns packet loss, min ping, average ping, and max ping
    time.sleep(1)
    pyautogui.typewrite(Device , interval=0) #Device
    pyautogui.hotkey('shift','tab')
    time.sleep(1)
    pyautogui.typewrite(LOCATION , interval=0) #Location
    pyautogui.hotkey('shift','tab')
    time.sleep(1)
    pyautogui.typewrite(data[data.index(":")+1:], interval=0) #Max Ping
    pyautogui.hotkey('shift','tab')
    time.sleep(1)
    pyautogui.typewrite(data[data.index("-")+1:data.index(":")], interval=0) #Avg Ping
    pyautogui.hotkey('shift','tab')
    time.sleep(1)
    pyautogui.typewrite(data[data.index("+")+1:data.index("-")], interval=0) #Min Ping
    pyautogui.hotkey('shift','tab')
    time.sleep(1)
    PL = data[:data.index("+")]
    pyautogui.typewrite(PL, interval=0) #Packet Loss
    pyautogui.hotkey('shift','tab')
    time.sleep(1)
    pyautogui.typewrite(str(datetime.datetime.now()), interval=0) #Date/Time

    pyautogui.typewrite(['enter'], interval=0)
   
    PL = data[:data.index("+")]
    return PL #returns packet loss for our message system

#-----------Main-----------
   
openUp()
time.sleep(5)
settingUp()

for k in range(NumOfDays): #This for loop gives us the option to run the system over multiple days without running overnight
    #time.sleep(72000) <---- approximately the time between when we want to run the system
    #If using this multi-day feature with the time.sleep(), make sure to be careful of when you start the system
    for i in range(NumOfRows):
        packetLoss = eachRow() # Runs each row and sets packetLoss equal to what ever it returns which could be a decimal or 'ERROR'
        if packetLoss != "ERROR" and float(packetLoss) <= 20.0: # Checks if the internet is doing well
            try:
                sendMessage(float(packetLoss), "*********")
                sendMessage(float(packetLoss))
               
                #sends messages to the team and Mr. C
                if float(packetLoss) > 2.5:
                    sendMessage(float(packetLoss), "**********
                   
                    #Only sends messages to Mr. McManly if the internet is doing abnormally poor
            except:
                print("Alert Failed")
        if packetLoss == "ERROR" or float(packetLoss) > 20.0: # Checks if the internet is down or doing very poorly
            try:
                sendMessage(100.0, "*********
                sendMessage(100.0)
                sendMessage(100.0, "*********
               
                # sends messages to everyone
            except:
                print("Alert Failed")
