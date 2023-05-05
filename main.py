# from cmath import log
import logging
# from socket import timeout
# from turtle import down
# from matplotlib.backend_bases import LocationEvent
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time,os,glob,logging,datetime,sys
from selenium.webdriver.firefox.options import Options
from datetime import date
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from bu_config import get_config
import bu_alerts
import sharepy

def remove_existing_files():
    #remove existing files from directory
    dir = os.getcwd() + '\\'+ 'temp'
    filelist = glob.glob(os.path.join(dir, "*"))
    for f in filelist:
        os.remove(f)
def login_and_download_process():
    try:
        profile = webdriver.FirefoxProfile()
        mime_types = ['application/pdf','Portable Document Format (PDF)', 'text/plain', 'application/vnd.ms-excel', 'text/csv', 'application/csv', 'text/comma-separated-values','application/x-pdf',
                'application/download', 'application/octet-stream', 'binary/octet-stream', 'application/binary', 'application/x-unknown']
        profile.set_preference("browser.helperApps.neverAsk.saveToDisk",
                        ",".join(mime_types))
        profile.set_preference("browser.download.manager.showWhenStarting", False)
        profile.set_preference("browser.download.dir", base_path)
        profile.set_preference("browser.download.folderList", 2)
        profile.set_preference('pdfjs.disabled',True)
        # profile.set_preference("browser.helperApps.neverAsk.openFile",",".join(mime_types))
        profile.set_preference("security.default_personal_cert", "Select Automatically")
        profile.set_preference("security.tls.version.min", 1)
        profile.accept_untrusted_certs = True

        binary = FirefoxBinary(r"C:\\Program Files\\Mozilla Firefox\\Firefox.exe")
        # binary = FirefoxBinary(r"C:\\Users\\chetan.surwade\\AppData\\Local\\Mozilla Firefox\\firefox.exe")
        driver = webdriver.Firefox(firefox_binary=binary, firefox_profile=profile,
                                    options=options, executable_path=executable_path)    
        driver.get('https://coderbyte.com/sl-org')
        time.sleep(5)
        WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="app"]/div[2]/div/div/div[1]/div[1]/div[1]/div/input'))).send_keys(username)
        time.sleep(5)
        WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div/div/div[1]/div[4]/button[1]'))).click()
        time.sleep(5)
        WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="app"]/div[2]/div/div/div[1]/div[1]/div[2]/div/input'))).send_keys(password)
        time.sleep(5)
        WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div/div/div[1]/div[4]/button[3]'))).click() # login button
        time.sleep(8)
        logging.info('Open plan and billing tab')
        try:
            logging.info('Open plan and billing tab')
            driver.get(url)
            time.sleep(7)
            driver.refresh()
            time.sleep(5)
        except:
            try:
                driver.get('https://coderbyte.com/dashboard/biourja-efzrr#settings-plan_and_billing')
                time.sleep(5)
                driver.refresh()
                time.sleep(120)
            except Exception as e:
                raise e
        check_file=WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/section[8]/div/div[3]/div[1]/div[2]/ul')))
        if check_file.text.split('\n')[1:][2]=='Upcoming invoice':
            month_file=check_file.text.split('\n')[1:][0].split(' ')[0]
            mail_subject_line=f'JOB SUCCESS - {job_name}. Upcoming file of {month_file} month still not came.'
            mail_body_line = f'{job_name} completed successfully, Attached PDF and logs.'
        else:
            new_file=WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/section[8]/div/div[3]/div[1]/div[2]/ul/li[2]/div[3]/a[2]')))
            new_file.click()   
            time.sleep(15)
            driver.switch_to.window(driver.window_handles[-1])
            WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div[1]/div/div[3]/div[1]/div/div[2]/table/tbody/tr[4]/td/div/button[1]/div/span'))).click()
            logging.info('Click on download button')
            time.sleep(15)
            logging.info('*****************Download Successfully*************')
            s = connect_to_sharepoint()
            file_upload_sp(s)
            logging.info("File is uploaded")
            logging.info("Sending mail for JOB SUCCESS")
            
            mail_subject_line=f'JOB SUCCESS - {job_name}. New File of {month_file} month uploaded into Sharepoint.'
            mail_body_line = f'{job_name} completed successfully, Attached PDF and logs \n {share_point_path}',

        logging.info("EMail Sent Successfully")
        bu_alerts.send_mail(
            receiver_email = receiver_email,
            mail_subject = mail_subject_line,
            mail_body = mail_body_line,
            multiple_attachment_list = locations_list
            )
        logging.info("EMail Sent Successfully")
    except Exception as e:
        logging.info(f"Exception caught during execution of login_and_download_process(): {e}")
        logging.exception(f'Exception caught during execution of login_and_download_process(): {e}')
        raise e
    finally:
        driver.quit()    


def connect_to_sharepoint():
    site = 'https://biourja.sharepoint.com'
    # Connecting to Sharepoint and downloading the file with sync params
    s = sharepy.connect(site, sp_username, sp_password)
    return s
    
        
# for upload the file on sharepoint
def file_upload_sp(s):
    filesToUpload = os.listdir(os.getcwd() + '\\'+ 'temp')
    for fileToUpload in filesToUpload:
        z=base_path+"\\"+fileToUpload
        locations_list.append(z)
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
        
def main():
    try:
        remove_existing_files()
        login_and_download_process()
    except Exception as e:
        logging.info(f"Exception caught in main(): {e}")
        logging.exception(f'Exception caught in main(): {e}')
        raise e

if __name__=="__main__" :
    try:
        logging.info('Execution Started')
        time_start = time.time()
        mydate = datetime.datetime.now()
        month = mydate.strftime("%b")
        year = datetime.date.today().year
        base_path = os.getcwd() + '\\'+'temp'
        options = Options()
        executable_path = os.getcwd()+"\\geckodriver_v0.30.exe"
        # logging 
        today_date=date.today()
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        # log file location
        logfile = os.getcwd() + '\\' + 'logs' + '\\' +'coderbyte_logfile'+str(today_date)+'.txt'
        if os.path.isfile(logfile):
            os.remove(logfile)

        logging.basicConfig(
            level=logging.INFO, 
            format='%(asctime)s [%(levelname)s] - %(message)s',
            filename=logfile)
        logging.info('Execution Started')

        locations_list=[]
        locations_list.append(logfile)

        # get credentials from bu_config
        credential_dict = get_config('CODER_BYTE_INVOICE_AUTOMATION','CODER_BYTE_INVOICE_AUTOMATION')
        job_name = credential_dict['PROJECT_NAME']
        username = credential_dict['USERNAME'].split(';')[0]
        password = credential_dict['PASSWORD'].split(';')[0]
        sp_username = credential_dict['USERNAME'].split(';')[1]
        sp_password =  credential_dict['PASSWORD'].split(';')[1]
        share_point_path = credential_dict['API_KEY']
        #receiver_email='enoch.benjamin@biourja.com'
        receiver_email = credential_dict['EMAIL_LIST']
        url=credential_dict['SOURCE_URL']
        main()
        time_end = time.time()
        logging.info(f"It takes {time_end-time_start} seconds to run")
        logging.info("Execution ended")
    except Exception as e:
        logging.info(f"Exception caught during execution: {e}")
        logging.exception(f'Exception caught during execution: {e}')
        bu_alerts.send_mail(
            receiver_email = receiver_email,
            mail_subject = 'JOB FAILED - CODER_BYTE_INVOICE_AUTOMATION',
            mail_body=f'CODER_BYTE_INVOICE_AUTOMATION failed during execution, Attached logs',
            attachment_location = locations_list
        )
        sys.exit(1)