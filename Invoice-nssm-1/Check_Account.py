from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.keys import Keys
from time import sleep
import os
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import urllib.request
from azcaptchaapi import AZCaptchaApi
import random
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
import aspose.words as aw
import json
import xml.etree.ElementTree as ET

# Parse the XML file
xml_file = 'Invoice_Setting.xml'
tree = ET.parse(xml_file)
root = tree.getroot()

variable = root.find("variable")
xpath_selector = root.find("xpath_selector")


#options = Options()
options = webdriver.ChromeOptions()
#options.add_experimental_option("detach",True)
options.add_argument("--headless")
options.add_argument("--window-size=1280,800");

options.add_argument("--disable-gpu")
#options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
#options.add_argument("--disable-extensions")
#options.add_experimental_option("useAutomationExtension", False)

# options.add_argument("--no-sandbox")
# options.add_argument("--allow-insecure-localhost")

#options.add_argument("start-maximized")

# browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
# browser = webdriver.Chrome("chromedriver.exe",options=options)


browser = webdriver.Chrome(options=options)
#browser.maximize_window()
API_KEY = variable.find("api_key").text
api = AZCaptchaApi(API_KEY)
uername = variable.find("uername").text
password = variable.find("password").text
input_svg = variable.find("input_svg").text
output_png = variable.find("output_png").text
web_link = variable.find("web_link").text
browser.get("{}".format(web_link))




def svg_to_png():
    doc = aw.Document()

    # create a document builder and initialize it with document object
    builder = aw.DocumentBuilder(doc)

    # insert SVG image to document
    shape = builder.insert_image(input_svg)

    # OPTIONAL
    # Calculate the maximum width and height and update page settings 
    # to crop the document to fit the size of the pictures.
    pageSetup = builder.page_setup
    pageSetup.page_width = shape.width
    pageSetup.page_height = shape.height
    pageSetup.top_margin = 0
    pageSetup.left_margin = 0
    pageSetup.bottom_margin = 0
    pageSetup.right_margin = 0

    # save as PNG
    doc.save(output_png)

def delay():
    sleep(random.randint(3,5))

status = "Sucess"
def solve_capcha():
    global status
    capcha = True
    while(capcha==True):
        try:
            src = browser.find_element('xpath',xpath_selector.find("capcha_img").text).get_attribute('src')
        except:
            capcha =False
            break
            
        try:
            img = urllib.request.urlretrieve(src, input_svg)[0]
        except ValueError:
            img = output_png
        svg_to_png()
        captcha = api.solve(output_png)
            
        # print('\nTry to get captcha answer')
        result = captcha.await_result()
            #print('\nGot answer: ' + result)
        try:
            result_captcha_box = WebDriverWait(browser, 2).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("capcha_box").text)))
        except:
            captcha=False
            break
        result_captcha_box.clear()
        result_captcha_box.send_keys(result.upper())
        result_captcha_box.send_keys(Keys.ENTER)
        try:
            pop_up = WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/span/div/div/div/div[1]')))
            if pop_up.text == "Tên đăng nhập hoặc mật khẩu không đúng":
                status =  "Wrong user name or password"
                capcha = False
                break
        except:
            pass
        try:
            WebDriverWait(browser, 3).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("capcha_img").text)))
            #print("\nwrong captcha")
            # áp dụng delay tránh trường hợp code chạy nhanh hơn quá trình reload hình ảnh captcha
            delay()
        except TimeoutException:
            captcha = False
        

def login():
    login_button = xpath_selector.find("login_button").text
    WebDriverWait(browser,60).until(EC.element_to_be_clickable((By.XPATH,"{}".format(login_button)))).click()

    user_name = WebDriverWait(browser,60).until(EC.element_to_be_clickable((By.XPATH,"{}".format(xpath_selector.find("user_name_box").text))))
    user_name.click()
    user_name.send_keys(uername)
    pass_word = WebDriverWait(browser,60).until(EC.element_to_be_clickable((By.XPATH,"{}".format(xpath_selector.find("password_box").text))))
    pass_word.click()
    pass_word.send_keys(password)
    print("Solve capcha")
    solve_capcha()
    
# print("login")
login()
with open("Check_Account.txt", "w") as file:
    file.write(status)
# print("Login sucessful")

browser.quit()











