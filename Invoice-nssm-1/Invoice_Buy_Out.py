
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
#options.add_argument("--headless")
options.add_argument("--window-size=1280,800");

#options.add_argument("--disable-gpu")
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

def solve_capcha():
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
    
print("login")
login()
print("Login sucessful")

invoice_search_link = variable.find("invoice_search_link").text
browser.get("{}".format(invoice_search_link))

def redirect_1():
    # buy_out_tab = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("buy_out_tab").text)))
    # buy_out_tab.click()
    search_button = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("search").text)))
    search_button.click()
    
invoices = {
    "invoices": []
}
    
def select_information_1():
    print("Redirect")
    redirect_1()
    #wait for table 
    WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("wait").text)))
    rows = browser.find_elements('xpath',xpath_selector.find("invoices_list").text)
    invoice_count=1
    for row in rows:
        invoice = {}
        row.click()
        view_button = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("view").text)))
        view_button.click()
        #wait for pop-up
        sleep(2)
        WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("wait_pop_up").text)))
        print("Get Invoice:{}".format(invoice_count))
        invoice["Mẫu_số"] =  browser.find_element('xpath',xpath_selector.find("Mau_so").text).text[7:]
        
        #Ki_hieu
        invoice["{}".format(browser.find_element(By.XPATH,xpath_selector.find("Ki_hieu").text).text[:7])] = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("Ki_hieu").text))).text[9:]
        #Số              
        invoice["{}".format(WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("So").text))).text[:2])] = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("So").text))).text[4:]
        #title
        invoice["Tên hóa đơn"] = browser.find_element('xpath',xpath_selector.find("title").text).text
        
        #ngày
        invoice["Ngày thành lập"] = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("date").text))).text.replace('\n',' ')
        #MCCQT
        invoice["{}".format(WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("MCCQT").find("key").text))).text)] = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("MCCQT").find("value").text))).text

        di_label=browser.find_elements('xpath',xpath_selector.find("di_label").text)
        di_value=browser.find_elements('xpath',xpath_selector.find("di_value").text)
       
        i=0
        for di in di_label:
            invoice["{}".format(di.text)] = di_value[i].text
            i+=1
                

        table_1 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("table_1").text)))
        title_table_1 = table_1.find_elements('xpath','.//th')
        row_table_1 = table_1.find_elements('xpath','.//tr')
        table_json_1 = {
            "rows":[]
        }
        i=0
        for row in row_table_1:
            row_json={}
            if(i>0):
                tds = row.find_elements('xpath','.//td')
                j=0
                for td in tds:
                    row_json["{}".format(title_table_1[j].text)] = td.text
                    j+=1
                table_json_1['rows'].append(row_json)
            i+=1
            
        table_json_2 = {}
        table_2 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("table_2").text)))
        title_table_2 = table_2.find_elements('xpath','.//th')
        row_table_2 = table_2.find_elements('xpath','.//tr')
        i=0
        for row in row_table_2:
            if(i>0):
                tds = row.find_elements('xpath','.//td')
                j=0
                for td in tds:
                    #print("{} : {}".format(title_table_2[j].text,td.text))
                    table_json_2["{}".format(title_table_2[j].text)] = td.text
                    j+=1
            i+=1
            
        table_json_3 = {}
        table_3 = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("table_3").text)))
        row_table_3 =table_3.find_elements('xpath','.//tr')
        for row in row_table_3:
            tds = row.find_elements('xpath','.//td')
            #print('{}:{}'.format(tds[0].text,tds[1].text))
            table_json_3["{}".format(tds[0].text)] = tds[1].text


        invoice["tables"] = [table_json_1,table_json_2,table_json_3]
        invoices["invoices"].append(invoice)
        invoice_count+=1
        exit_button = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,xpath_selector.find("exit_button").text)))
        exit_button.click()
        sleep(1)
        
json_file = variable.find("json_out").text
def export_json_1():
    
    with open(json_file,"w",encoding='utf-8') as f:
        json.dump(invoices,f, ensure_ascii=False, indent=2)
select_information_1()
print("Export data to json")
export_json_1()

#sleep(50)
browser.quit()