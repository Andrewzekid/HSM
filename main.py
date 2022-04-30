import pyautogui
import keyboard
import smtplib
import time
import random
import win32api
from email.message import EmailMessage
from email.header import Header

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



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
    time.sleep(3)

def locate_and_click(path,wait_before_start=0):
    if wait_before_start == 0:
            x,y = pyautogui.locateCenterOnScreen(path,confidence=0.7)
            pyautogui.moveTo(x,y,duration=0.5,tween=pyautogui.easeInQuad)
            pyautogui.click(x,y)
    else:
            time.sleep(wait_before_start)
            x,y = pyautogui.locateCenterOnScreen(path,confidence=0.7)
            pyautogui.moveTo(x,y,duration=0.5,tween=pyautogui.easeInQuad)
            pyautogui.click(x,y)


def smtp_alert(keyword,usern='andrewwang417@gmail.com',password='dhaxugvztawgjulf',to_address='yw3479@email.kist.ed.jp'):
    with open("template.txt","w",encoding="utf-8") as f:
        f.write("""
    {} has now been added to the Shopping Cart!
    Please check the computer immediatly.

    --Hermes Scraping Manager
    """.format(keyword))
        f.close()
    
    with open("template.txt","r",encoding="utf-8") as f:
        message = f.read()

    print(message)

    gmail_user = usern
    gmail_password = password

    sent_from = gmail_user
    to = to_address

    msg = EmailMessage()
    msg['Subject'] = '{} has now been added to the Shopping Cart!'.format(keyword)
    msg['To'] = to
    msg["From"] = sent_from
    
    msg.set_content(message)



    # msg.set_content("""
    # {} has now been added to the Shopping Cart!
    # Please check the computer immediatly.

    # --Hermes Scraping Manager
    # """.format(keyword),charset="utf-8")


   
    # subject = 'Please check the computer immediately. '
    # body = '{} has now been added to the Shopping Cart!'.format(keyword)

    # email_text = """\
    # From: %s
    # To: %s
    # Subject: %s

    # %s
    # """ % (sent_from, ", ".join(to), subject, body)

    try:
        smtp_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        smtp_server.ehlo()
        smtp_server.login(gmail_user, gmail_password)
        smtp_server.send_message(msg)
        smtp_server.close()
        print ("Email sent successfully!")
    except Exception as ex:
        print ("Something went wrong….",ex)

if __name__=="__main__":
    #set encoding method for japanese email
    #'leonaellen@gmail.com',

    access_website()
    #attempt to locate targeted bag
    while not(keyboard.is_pressed("esc")):
        try:
            locate_and_click("keyword.jpg")
            print("Success!")
        except:
            #refresh the page
            locate_and_click("refresh.jpg")
            time.sleep(1)
        else:
            #targeted bag has been found
            print("Target identified!")
            locate_and_click("keyword.jpg")
            time.sleep(1)
            locate_and_click("cross.png")
            time.sleep(1.2)
            locate_and_click("shopping_cart.jpg")
            time.sleep(3)
            locate_and_click("viewshoppingbag.JPG")
            smtp_alert(keyword=" クラッチバッグ 《ヴェルー》 ")
            break