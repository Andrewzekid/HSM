from numpy import number
import pyautogui
import keyboard
import smtplib
import time
import random
import win32api
from email.message import EmailMessage
from email.header import Header
import math
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



def access_website(wait_before_start=20):
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
    time.sleep(wait_before_start)

def locate_and_click(path,wait_before_start=0,duration=0.5,grayscale=False,tween=pyautogui.easeInQuad,confidence=0.7):
    if wait_before_start == 0:
            x,y = pyautogui.locateCenterOnScreen(path,confidence=confidence,grayscale=grayscale)
            pyautogui.moveTo(x,y,duration=duration,tween=tween)
            pyautogui.click(x,y)
    else:
            time.sleep(wait_before_start)
            x,y = pyautogui.locateCenterOnScreen(path,confidence=confidence)
            pyautogui.moveTo(x,y,duration=duration,tween=tween)
            pyautogui.click(x,y)

def zoomOut(optionspath,number_of_zooms):
    """
    zoomOut: A function designed to automate zooming out

    Parameters:
        optionspath: the path to the image of the options icon
        number_of_zooms: number of times you want to click on the zoomout button
    """
    locate_and_click(optionspath,grayscale=True,duration=0.01,tween=None)
    time.sleep(0.25)

    x2,y2 = pyautogui.locateCenterOnScreen("zoomout.jpg",confidence=0.7,grayscale=True)
    for _ in range(number_of_zooms):
        pyautogui.click(x2,y2)

def smtp_alert(keyword,usern='andrewwang417@gmail.com',password='dhaxugvztawgjulf',to_address='andrewwang348@gmail.com'):
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

def login():
    #click on the account button
    locate_and_click("account.jpg",duration=0.01,tween=None)

    # #wait until the login button is located
    x,y = wait_until_image_located_on_screen("login.jpg",duration=4)
    pyautogui.moveTo(x,y)
    pyautogui.click(x,y)
    
    x2,y2 = wait_until_image_located_on_screen("logincross.jpg",duration=30)
    pyautogui.click(x2,y2)
    time.sleep(.3)

def wait_until_image_located_on_screen(path,duration,confidence=0.7,grayscale=False):
    #calculate the number of times to try looping
    no_loops = math.floor(duration/0.1)

    for _ in range(no_loops):
        try:
            x,y = pyautogui.locateCenterOnScreen(path,confidence=confidence,grayscale=grayscale)
            time.sleep(0.1)
        except:
            #cant find required button
            pass
        else:
            print(x,y)
            return x,y

def refresh(path):
    x,y = wait_until_image_located_on_screen("refresh.jpg",duration=8)
    pyautogui.click(x,y)
    time.sleep(.3)

def make_payment(creditcardname,creditcardnumber,creditcardexpirydate,securitycode):
    x4,y4 = wait_until_image_located_on_screen("continuepayment.jpg",duration=20)
    pyautogui.click(x4,y4)

    zoomOut("optionsdarkblue.jpg",number_of_zooms=4)

    x2,y2 = wait_until_image_located_on_screen("continue.jpg",duration=20)
    pyautogui.click(x2,y2)
    
    x3,y3 = wait_until_image_located_on_screen("creditcardbutton.jpg",confidence=0.95,duration=10)
    pyautogui.click(x3,y3)
    
    #pause for a bit to allow the full form to load
    time.sleep(0.75)

    #click the input button that allows you to enter your name
    locate_and_click("creditcardname.jpg",duration=0.01,tween=None)
    pyautogui.typewrite(creditcardname)

    #click the input button that allows you to enter your credit card number
    locate_and_click("creditcardnumber.jpg",duration=0.01,tween=None)
    pyautogui.typewrite(creditcardnumber)

    #click the input button that allows you to enter the expiry date
    locate_and_click("expirydate.jpg",duration=0.01,tween=None)
    pyautogui.typewrite(creditcardexpirydate)
    
    #click the input button that allows you to enter your CV number
    locate_and_click("securitycode.jpg",duration=0.01,tween=None)
    pyautogui.typewrite(securitycode)
    
    #agree to the terms and conditions
    locate_and_click("agreetoterms.jpg",duration=0.01,tween=None,confidence=0.95)

    #continue with payment
    # locate_and_click("pay.jpg")
    # time.sleep(0.5)


if __name__=="__main__":
    #set encoding method for japanese email
    #'leonaellen@gmail.com',
    # time.sleep(3)
    # zoomOut("optionsdarkblue.jpg",number_of_zooms=4)
    access_website()
    login()

    #attempt to locate targeted bag
    while not(keyboard.is_pressed("esc")):
        try:
            locate_and_click("keyword.jpg",grayscale=True)
            print("Success!")
        except:
            #refresh the page
            refresh("refresh.jpg")
        else:
            #targeted bag has been found
            print("Target identified!")
            locate_and_click("keyword.jpg",grayscale=True)
            time.sleep(1)
            try:
                locate_and_click("cross.png",grayscale=True)
                time.sleep(0.2)
            except:
                pass

            locate_and_click("shopping_cart.jpg",grayscale=True)
            
            #wait until the product has been added to the shopping cart
            x1,y1 = wait_until_image_located_on_screen("viewshoppingbag.JPG",duration=20,grayscale=True)
            pyautogui.click(x1,y1)

            make_payment("hello world","1234","05/2022","234")

            # smtp_alert(keyword=" クラッチバッグ 《ヴェルー》 ")
            # break