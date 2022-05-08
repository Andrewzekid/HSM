import win32api
import time
import pyautogui
import keyboard
import pyautogui
import keyboard
import smtplib
import time
import win32api
from email.message import EmailMessage
import math
from playsound import playsound
from time import perf_counter

class MultiSelect:
    def __init__(self,target,*args,**kwargs):
        """
        Description:
            Initialization function for Multiselect object

        Parameters:
            target[list]: a list of tuples in (productname,pathtoimageofproductname) format.
        """
        self.target = target
    
    def search(self):
        """
        Description:
            A search function that allows for the snatching of multiple different types of bags
        """
        found = False
        while not(keyboard.is_pressed("esc")):
            for productname,pathtoimageofproductname in self.target:
                #loop through the target list
                try:
                    locate_and_click(pathtoimageofproductname,grayscale=True)
                except:
                    pass
                else:
                    t1_start = perf_counter()
                    #this part will run if the locate and click is successful
                    print("Target identified!")

                    #targeted bag has been found
                    x3,y3 = wait_until_image_located_on_screen("images/shopping_cart.jpg",duration=10,grayscale=True)
                    pyautogui.moveTo(x3,y3,duration=0.3,tween=pyautogui.easeInOutBounce)
                    pyautogui.click(x3,y3)
                    
                    #wait until the product has been added to the shopping cart
                    x1,y1 = wait_until_image_located_on_screen("images/viewshoppingbag.JPG",duration=20,grayscale=True)
                    pyautogui.moveTo(x3,y3,duration=0.3,tween=pyautogui.easeInOutBounce)
                    pyautogui.click(x1,y1)

                    #creditcardname,creditcardnumber,creditcardexpirydate,securitycode
                    make_payment("KA OU","4297 7100 0065 6460","04/27","580")
                    
                    t1_stop = perf_counter()
                    print("Elapsed time:", t1_stop, t1_start)
                    print("Elapsed time during the whole program in seconds:",
                                        t1_stop-t1_start)    

                    playsound("music/Cornered.mp3")

                    smtp_alert(keyword=productname)
                    found = True       
                    break
            
            else:
                if found:
                    break
                else:
                    refresh("images/refresh.jpg")

def use_vpn():
    locate_and_click("images/extension.jpg")
    time.sleep(0.5)
    locate_and_click("images/touchvpnextension.jpg")
    time.sleep(1)
    locate_and_click("images/connect.jpg")
    time.sleep(2)
                    
def access_website(wait_before_start=20):
    win32api.LoadKeyboardLayout('00000409',1) # to switch to english
    x,y = pyautogui.locateCenterOnScreen("images/chrome.jpg",confidence=0.7)
    pyautogui.moveTo(x,y)
    pyautogui.click(x,y)

    time.sleep(1)

    x2,y2 = pyautogui.locateCenterOnScreen("images/chromeprofile.jpg",confidence=0.7)
    pyautogui.moveTo(x2,y2)
    pyautogui.click(x2,y2)

    time.sleep(0.5)
    use_vpn()

    with pyautogui.hold("ctrl"):
        pyautogui.press("t")
    
    pyautogui.typewrite("https://www.hermes.com/jp/ja/category/women/bags-and-small-leather-goods/bags-and-clutches/")
    pyautogui.press("enter")
    time.sleep(wait_before_start)

def locate_and_click(path,wait_before_start=0,duration=0.3,grayscale=False,tween=pyautogui.easeInQuad,confidence=0.7):
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

    x2,y2 = pyautogui.locateCenterOnScreen("images/zoomout.jpg",confidence=0.7,grayscale=True)
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
    locate_and_click("images/account.jpg",duration=0.01,tween=None)

    # #wait until the login button is located
    x,y = wait_until_image_located_on_screen("images/login.jpg",duration=4)
    pyautogui.moveTo(x,y)
    pyautogui.click(x,y)
    
    time.sleep(5)

    x2,y2 = wait_until_image_located_on_screen("images/logincross.jpg",duration=30)
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
    x,y = wait_until_image_located_on_screen("images/refresh.jpg",grayscale=True,duration=8)
    pyautogui.click(x,y)
    time.sleep(3)

def make_payment(creditcardname,creditcardnumber,creditcardexpirydate,securitycode):
    x4,y4 = wait_until_image_located_on_screen("images/continuepayment.jpg",grayscale=True,duration=10)
    pyautogui.click(x4,y4)

    zoomOut("images/optionsdarkblue.jpg",number_of_zooms=4)

    x2,y2 = wait_until_image_located_on_screen("images/continue.jpg",grayscale=True,duration=20)
    pyautogui.click(x2,y2)
    
    x3,y3 = wait_until_image_located_on_screen("images/creditcardbutton.jpg",grayscale=True,confidence=0.95,duration=10)
    pyautogui.click(x3,y3)
    
    #pause for a bit to allow the full form to load
    time.sleep(0.75)

    #click the input button that allows you to enter your name
    locate_and_click("images/creditcardname.jpg",grayscale=True,duration=0.01,tween=None)
    pyautogui.typewrite(creditcardname)

    #click the input button that allows you to enter your credit card number
    locate_and_click("images/creditcardnumber.jpg",grayscale=True,duration=0.01,tween=None)
    pyautogui.typewrite(creditcardnumber)

    #click the input button that allows you to enter the expiry date
    locate_and_click("images/expirydate.jpg",grayscale=True,duration=0.01,tween=None)
    pyautogui.typewrite(creditcardexpirydate)
    
    #click the input button that allows you to enter your CV number
    locate_and_click("images/securitycode.jpg",grayscale=True,duration=0.01,tween=None)
    pyautogui.typewrite(securitycode)
    
    #agree to the terms and conditions
    locate_and_click("images/agreetoterms.jpg",grayscale=True,duration=0.01,tween=None,confidence=0.95)

    #continue with payment
    # locate_and_click("pay.jpg")
    # time.sleep(0.5)