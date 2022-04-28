import pyautogui
import keyboard
import smtplib
import time
import random
import win32api
from email.message import Message
from email.header import Header

def access_website():
    win32api.LoadKeyboardLayout('00000409',1) # to switch to english
    x,y = pyautogui.locateCenterOnScreen("chrome.jpg",confidence=0.7)
    pyautogui.moveTo(x,y)
    pyautogui.click(x,y)

    time.sleep(1)

    x2,y2 = pyautogui.locateCenterOnScreen("chromeprofile.jpg",confidence=0.7)
    pyautogui.moveTo(x2,y2)
    pyautogui.click(x2,y2)

    time.sleep(0.5)

    with pyautogui.hold("ctrl"):
        pyautogui.press("t")
    
    pyautogui.typewrite("https://www.hermes.com/jp/ja/category/women/bags-and-small-leather-goods/bags-and-clutches/")
    pyautogui.press("enter")
    time.sleep(random.randint(1,3))

def locate_and_click(path,wait_before_start=0):
    if wait_before_start == 0:
        x,y = pyautogui.locateCenterOnScreen(path,confidence=0.7)
        pyautogui.moveTo(x,y,duration=0.5,tween=pyautogui.easeInElastic)
        pyautogui.click(x,y)
    else:
        time.sleep(wait_before_start)
        x,y = pyautogui.locateCenterOnScreen(path,confidence=0.7)
        pyautogui.moveTo(x,y,duration=0.5,tween=pyautogui.easeInElastic)
        pyautogui.click(x,y)

def smtp_alert(keyword,usern='andrewwang417@gmail.com',password='dhaxugvztawgjulf',to_address=['yw3479@email.kist.ed.jp']):
    keyword = keyword.encode(encoding="UTF-8",errors="strict")
    gmail_user = usern
    gmail_password = password

    msg = Message()
    msg['Subject'] = Header('{} has now been added to the Shopping Cart!'.format(keyword), 'utf-8')


    sent_from = gmail_user
    to = to_address
    subject = 'Please check the computer immediately. '
    body = '{} has now been added to the Shopping Cart!'.format(keyword)

    email_text = """\
    From: %s
    To: %s
    Subject: %s

    %s
    """ % (sent_from, ", ".join(to), subject, body)

    try:
        smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        smtp_server.ehlo()
        smtp_server.login(gmail_user, gmail_password)
        smtp_server.sendmail(sent_from, to, email_text)
        smtp_server.close()
        print ("Email sent successfully!")
    except Exception as ex:
        print ("Something went wrong….",ex)

if __name__=="__main__":
    #set encoding method for japanese email
    #'leonaellen@gmail.com',

    # access_website()
    smtp_alert(keyword=" クラッチバッグ 《ヴェルー》 ")