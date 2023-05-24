from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import csv
import os
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException
import pandas as pd
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
import getpass
import warnings
import numpy as np
import unidecode
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options  = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920x1080") 
    path = ChromeDriverManager().install()
    driver = webdriver.Chrome(path, options=chrome_options)
    driver.maximize_window()

    return driver, path, chrome_options

def login_Artworks(driver, name, pwd):
    URL1 = "https://masterworks.io"
    # navigating to the website link
    driver.get(URL1)
    time.sleep(3)
    # closing welcome window
    try:
        button = driver.find_element_by_xpath("//button[@type='button' and @class='gTx50DUf ySKBdvn6 apL75T7Q ZOT6B5xU v3COHnQD n0JHHCrz Qe1jFcyE']")
        driver.execute_script("arguments[0].click();", button)
    except:
        pass

    time.sleep(2)    
    buttons = driver.find_elements_by_tag_name("button")
    login = False
    for button in buttons:
        if button.text.lower() == 'sign in':
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
            login = True
            break

    if not login:
        return

    # signing in 
    print('Signing In ...')
    print('-'*75)
    username = driver.find_element_by_name("username")
    password = driver.find_element_by_name("password")
    username.send_keys(name)
    time.sleep(1)
    password.send_keys(pwd)
    time.sleep(1)
    button = driver.find_element_by_xpath("//input[@type='submit' and @value='Sign in']")
    driver.execute_script("arguments[0].click();", button)
    time.sleep(2)

def scrape_portfolio(driver):
   
    # navigating to movements web page
    driver.get('https://www.masterworks.io/dashboard/portfolio')
    time.sleep(3)
    buttons = driver.find_elements_by_tag_name("button")
    for button in buttons:
        if button.text.lower() == 'masterworks portfolio values':
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
            break
    # Scraping table data
    rows = []
    html = driver.page_source
    soup = BeautifulSoup(html, 'lxml')
    items = soup.find('table').find("tbody").find_all("tr")
    for item in items:
        info = item.find_all(['td'])
        inds = [1, 2, 3, 5, 7, 8]
        row = []
        for i, data in enumerate(info):
            if i in inds:
                text = str(data.text)
                text = text.split(',')[0]
                text = text.replace("—", "").replace('*', '').replace('Early Days', '').replace('$', '')
                text = unidecode.unidecode(text)
                row.append(text)

        rows.append(row.copy())

    header = ['Investment Name', 'Artist', 'Artwork', 'Est. NAV Per Share', "Implied Gross Change", "Last Appraisal Date"]

    file1 = 'portfolio.csv'
    path = os.getcwd() + '\\' + file1
    if os.path.isfile(path):
        os.remove(path) 
        #time.sleep(10)

    with open(file1, 'w', newline='') as file:
        csvwriter = csv.writer(file)
        csvwriter.writerow(header)
        csvwriter.writerows(rows)

def scrape_sell_orders(driver):
   
    driver.get('https://www.masterworks.io/trading/bulletin')
    time.sleep(3)
    start = False
    rows = []
    while True:
        # Scraping table data
        html = driver.page_source
        soup = BeautifulSoup(html, 'lxml')
        items = soup.find_all('table')[0].find("tbody").find_all("tr")
        for item in items:
            cells = item.find_all('td')
            row = []
            for i, data in enumerate(cells):
                # scraping the first four columns only
                if i == 4: break
                # numerical columns scraping
                if not data.find('span'):
                    text = data.text
                    text = text.replace('*', '').replace('$', '')
                    text = unidecode.unidecode(text)
                    row.append(text) 
                # Artwork column scraping
                else:
                    words = data.find_all('span')
                    for word in words:
                        text = word.text
                        text = text.replace('*', '').replace('$', '')
                        text = unidecode.unidecode(text)
                        # Artwork data
                        if text.count('Masterworks') > 0:
                            row.append(text[:-17])
                        # Artist data
                        else:
                            row.append(text)
                        # investment name data (e.g. Masterworks 001)
                        if word.find('small'):
                            text = word.find('small').text
                            text = unidecode.unidecode(text)
                            row.append(text[1:-1])

            rows.append(row.copy())
        try:
            elements = driver.find_elements_by_xpath("//button[@type='button' and @class='gTx50DUf OjHzwgqq ySKBdvn6 apL75T7Q v3COHnQD MPlI1aj4 Qe1jFcyE']")
            if len(elements) == 2 and not start:
                button = elements[0]
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
                start = True
            elif len(elements) == 3:
                button = elements[1]
                driver.execute_script("arguments[0].click();", button)
                time.sleep(1)
            else:
                break
        except:
            break

    header = ['Artist', 'Artwork', 'Investment Name', 'Total', 'Quantity', 'Ask']

    file1 = 'sell_prices.csv'
    path = os.getcwd() + '\\' + file1
    if os.path.isfile(path):
        os.remove(path) 

    with open(file1, 'w', newline='') as file:
        csvwriter = csv.writer(file)
        csvwriter.writerow(header)
        csvwriter.writerows(rows)

def processing_data():

    files = ['portfolio.csv', 'sell_prices.csv']
    df_port = pd.read_csv(os.getcwd() + '\\' + files[0])
    df_sell = pd.read_csv(os.getcwd() + '\\' + files[1])
    df_out = df_port.merge(df_sell, how='outer', on='Investment Name')
    df_out.drop(['Artist_y', 'Artwork_y'], inplace=True, axis=1)
    df_out.rename(columns = {'Artist_x':'Artist', 'Artwork_x': 'Artwork'}, inplace=True)
    df_out.dropna(subset=['Ask'], inplace=True)
    df_out.fillna(0, inplace=True)
    df_out[['Ask', 'Est. NAV Per Share']] = df_out[['Ask', 'Est. NAV Per Share']].astype(str)
    df_out['Ask'] = df_out['Ask'].apply(lambda x: x.replace(',', ''))
    df_out['Est. NAV Per Share'] = df_out['Est. NAV Per Share'].apply(lambda x: x.replace(',', ''))
    df_out[['Ask', 'Est. NAV Per Share']] = df_out[['Ask', 'Est. NAV Per Share']].astype(float)
    df_out['Variation %'] = ((df_out['Est. NAV Per Share'] - df_out['Ask'])/df_out['Ask'])*100
    df_out.sort_values('Variation %', inplace=True, ascending=False)
    df_out.to_csv('scraping_results.csv', encoding='UTF-8', index=False)

    for file in files:
        path = os.getcwd() + '\\' + file
        if os.path.isfile(path):
            os.remove(path) 

    max_var = np.max(df_out['Variation %'].values)
    #print(max_var)
    return max_var

def clear_screen():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
  
    # for mac and linux
    else:
        _ = os.system('clear')

def send_mail(sender_mail, sender_pass, recip_mail):

    fromaddr = sender_mail
    toaddr = recip_mail

    # MIMEMultipart 
    msg = MIMEMultipart() 
    # senders email address 
    msg['From'] = fromaddr 
    # receivers email address 
    msg['To'] = toaddr 
    # the subject of mail
    msg['Subject'] = "Variations Above 20% In Prices!"
    # the body of the mail 
    body = "Kindly check the attached sheet for all the items with price variations above 20%."
    # attaching the body with the msg 
    msg.attach(MIMEText(body, 'plain')) 
    # open the file to be sent
    # rb is a flag for readonly 
    filename = "scraping_results.csv"
    attachment = open(os.getcwd() + '\\' + filename, "rb") 
    # MIMEBase
    attac= MIMEBase('application', 'octet-stream') 
    # To change the payload into encoded form 
    attac.set_payload((attachment).read()) 
    # encode into base64 
    encoders.encode_base64(attac) 
    attac.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
    # attach the instance 'p' to instance 'msg' 
    msg.attach(attac) 
    # creates SMTP session 
    email = smtplib.SMTP('smtp.gmail.com', 587) 
    # TLS for security 
    email.starttls() 
    # authentication 
    email.login(fromaddr, sender_pass) 
    # Converts the Multipart msg into a string 
    message = msg.as_string() 
    # sending the mail 
    email.sendmail(fromaddr, toaddr, message) 
    # terminating the session 
    email.quit() 


# Website login credentials
name1 = ""
pwd1 = ""

os.system("color")
COLOR = {"HEADER": "\033[95m", "BLUE": "\033[94m", "GREEN": "\033[92m",
            "RED": "\033[91m", "ENDC": "\033[0m", "YELLOW" :'\033[33m'}

while(True):

    name1 = input('Please enter the website login username: ')
    if len(name1) == 0:
        print('Invalid username! Please try again.')
        continue
    pwd1 = getpass.getpass('Please enter the website login password: ')
    if len(pwd1) == 0:
        print('Invalid password! Please try again.')
        continue

    break

 #update rate for the data
while(True):
    try:
        user_input = input('Please enter the time interval (in minutes) for updating the data: ')
        user_input = int(user_input)
    except:
        print('Invalid Input! Please try again.')
        continue

    if user_input  > 0:
        break
    else:
        print('Invalid Input! Please try again.')

sender_mail = ''
sender_pass = ''
recip_mail = ''

while(True):

    sender_mail = input('Please enter the sender mail address: ')
    if len(sender_mail) > 0:
        if sender_mail.count('@') == 0:
            print('Invalid mail address! Please try again.')
            continue
        sender_pass = getpass.getpass('Password: ')
    else:
        sender_mail = 'scraper.notification1@gmail.com'
        sender_pass = 'asd2376312'
    break
        
while(True):

    recip_mail = input('Please enter the recipient mail address: ')
    if len(sender_mail) > 0:
        if sender_mail.count('@') == 0:
            print('Invalid mail address! Please try again.')
            continue
    break

update_rate = user_input * 60

if __name__ == '__main__':

    driver, path, options = initialize_bot()
    clear_screen()
    print('Accessing the website ...')
    print('-'*75)
    nfail = 0
    login = True
    while(True):
        try:
            if login:
                login_Artworks(driver, name1, pwd1)
                print('Fetching the data from the website ...')
                print('-'*75)
                login = False
            scrape_portfolio(driver)
            scrape_sell_orders(driver)
            max_var = processing_data()
            if max_var > 20.0:
                try:
                    send_mail(sender_mail, sender_pass, recip_mail)
                except:
                    row ='  Failure in sending the notification mail!'
                    print(COLOR["YELLOW"],row, COLOR["ENDC"])

            row = "  The data has been updated successfully in {}".format(datetime.now().strftime("%d/%m/%Y %H:%M"))
            print(COLOR["GREEN"],row, COLOR["ENDC"])
            time.sleep(update_rate)
            nfail = 0
        except:
            if nfail < 10:
                nfail += 1
                row ='  Failure in accessing the websites! Re-attempt number {} ...'.format(nfail)
                print(COLOR["YELLOW"],row, COLOR["ENDC"])
                driver.close()
                driver = webdriver.Chrome(path, options=options)
                time.sleep(60)
                login = True
            else:
                row ='  Failure in accessing the websites after 10 attemps! Exiting the program...'
                print(COLOR["RED"],row, COLOR["ENDC"])
                exit()