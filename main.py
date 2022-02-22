from cmath import log
from configparser import SectionProxy
from genericpath import exists
from lib2to3.pgen2 import driver
from logging import exception
import logging
from socket import timeout
from turtle import down
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
import time,os,glob,logging,datetime,sys
from datetime import date
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from bu_config import get_config, SharePoint
import bu_alerts
import sharepy

mydate = datetime.datetime.now()
month = mydate.strftime("%b")
year = datetime.date.today().year

# Credentials --
# USER = "pooja.upadhyay@biourja.com"
# PASSWORD = "Indore@123"

# logging 
today_date=date.today()
logfile =  os.getcwd() +'\\logs\\'+'coderbyte_logfile'+str(today_date)+'.txt'
logging.basicConfig(level=logging.INFO,filename=logfile,filemode='w',format='[line :- %(lineno)d] %(asctime)s [%(levelname)s] - %(message)s ')


profile = webdriver.FirefoxProfile()
profile.set_preference("browser.helperApps.neverAsk.saveToDisk","text/plain,application/pdf")
profile.set_preference("browser.download.manager.showWhenStarting",False)
profile.set_preference('browser.download.dir',os.getcwd() + '\\'+'temp')
profile.set_preference("browser.download.folderList",2)
profile.set_preference('pdfjs.disabled',True)
driver = webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)

# get credentials from bu_config
credential_dict = get_config('CODER_BYTE_INVOICE_AUTOMATION','CODER_BYTE_INVOICE_AUTOMATION')
username = credential_dict['USERNAME'].split(';')[0]
password = credential_dict['PASSWORD'].split(';')[0]
sp_username = credential_dict['USERNAME'].split(';')[1]
sp_password =  credential_dict['PASSWORD'].split(';')[1]
share_point_path = credential_dict['API_KEY']
receiver_email = credential_dict['EMAIL_LIST']

job_name = 'CODER_BYTE_INVOICE_AUTOMATION'

def connect_to_sharepoint():
    site = 'https://biourja.sharepoint.com'
    # Connecting to Sharepoint and downloading the file with sync params
    s = sharepy.connect(site, sp_username, sp_password)
    return s
    
        
# for upload the file on sharepoint
def file_upload_sp(s):
    filesToUpload = os.listdir(os.getcwd() + '\\'+ 'temp')
    for fileToUpload in filesToUpload:
        headers = {"accept": "application/json;odata=verbose",
        "content-type": "Portable Document Format (PDF)"}
        with open(os.path.join(os.getcwd() + '\\'+'temp', f"{fileToUpload}"), 'rb') as read_file:

            content = read_file.read()
        #  s.post(site + path1 + path2.format("/add(url='"+file+"',overwrite=true)"), data=content, headers=headers)
        # site =  'https://biourja.sharepoint.com'
        # path1 = '/BiourjaPower/_api/web/GetFolderByServerRelativeUrl'
        p = s.post(f"https://biourja.sharepoint.com/BiourjaPower/_api/web/GetFolderByServerRelativeUrl('Shared%20Documents/Power%20Reference/Power_Invoices/Coderbyte/')/Files/add(url='Coderbyte-{month} {year} Invoice.pdf',overwrite=true)",data=content,headers=headers)
        # p = s.post(f"{site}{path1}('Shared{share_point_path}')/Files/add(url='Coderbyte-{month} {year} Invoice.pdf',overwrite=true)",data=content,headers=headers)
    return p


# this function will wait until the file download then close the browser 
def download_wait(path_to_downloads= os.getcwd() + '\\temp'):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 90:
        time.sleep(1)
        dl_wait = True
        for fname in os.listdir(path_to_downloads):
            if fname.endswith('.pdf'):
                dl_wait = False
        seconds += 1
    time.sleep(seconds)
    driver.quit()
    return seconds
        
def main():
    try:
        # remove existing files from directory
        dir = os.getcwd() + '\\'+ 'temp'
        filelist = glob.glob(os.path.join(dir, "*"))
        for f in filelist:
            os.remove(f)
        
        logging.info('Open Coderbyte in firefox')
        driver.get('https://coderbyte.com/sl-org')         #login url
        logging.info('pass user name')
        driver.find_element(By.XPATH,'//*[@id="app"]/div[2]/div/div/div[1]/div[1]/div[1]/div/input').send_keys(username)
        logging.info('pass password')
        driver.find_element(By.XPATH,'//*[@id="app"]/div[2]/div/div/div[1]/div[1]/div[2]/div/input').send_keys(password)
        logging.info('click on Login Button')
        driver.find_element(By.XPATH,'//*[@id="app"]/div[2]/div/div/div[1]/div[2]/button').click() # login button
        time.sleep(8)
        
        try:
            logging.info('Open plan and billing tab')
            driver.get('https://coderbyte.com/dashboard/biourja-efzrr#settings-plan_and_billing')
            time.sleep(2)
            driver.refresh()
            WebDriverWait(driver,30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/section[8]/div/div[3]/div[2]/div/ul/li[2]/div[1]/a/span')))
            
        except:
            # driver.refresh()
            driver.get('https://coderbyte.com/dashboard/biourja-efzrr#settings-plan_and_billing')
            WebDriverWait(driver,30).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/section[8]/div/div[3]/div[2]/div/ul/li[2]/div[1]/a/span')))

        if month == driver.find_element(By.XPATH,'/html/body/div[1]/section[8]/div/div[3]/div[2]/div/ul/li[2]/div[1]/a/span').text.split()[0]:
            driver.find_element(By.XPATH,'/html/body/div[1]/section[8]/div/div[3]/div[2]/div/ul/li[2]/div[1]/a/span').click()  #latest invoice
        else:
            logging.info("Sending mail for JOB FAILED")
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject ='JOB FAILED - {}'.format(job_name),mail_body = 'No new invoice found for {}'.format(month),attachment_location = logfile)
            sys.exit() 
            
        driver.switch_to.window(driver.window_handles[-1])
        WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div[1]/div/div[3]/div[1]/div/div[2]/table/tbody/tr[4]/td/div/button[1]/div/span')))
        logging.info('Click on download button')
        driver.find_element_by_xpath('/html/body/div/div/div[1]/div/div[3]/div[1]/div/div[2]/table/tbody/tr[4]/td/div/button[1]/div/span').click()  #download invoice
        # time.sleep(10)
        logging.info('*****************Download Successfully*************')
        logging.info(f'File is downloaded in {download_wait()} seconds.')
        s = connect_to_sharepoint()
        file_upload_sp(s)
        logging.info("File is uploaded")
        logging.info("Sending mail for JOB SUCCESS")
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully, Attached logs \n Click on the given link for latest INVOICE \n {share_point_path} ',attachment_location = logfile)
        logging.info("EMail Sent Successfully")
        
    except Exception as e:
        logging.exception(str(e))
        logging.info("Sending mail for JOB FAILED")
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject ='JOB FAILED - {}'.format(job_name),mail_body = 'Error in main() details {}'.format(str(e)),attachment_location = logfile)
    # finally:
    #     driver.quit()

if __name__=="__main__" :
    logging.info('Execution Started')
    time_start = time.time()
    
    main()
    
    time_end = time.time()
    logging.info(f"It takes {time_end-time_start} seconds to run")
    logging.info("Execution ended")