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
import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.encoders as encoders


mydate = datetime.datetime.now()
month = mydate.strftime("%b")
year = datetime.date.today().year
base_path = os.getcwd() + '\\'+'temp'
options = Options()
# Credentials --
# USER = "pooja.upadhyay@biourja.com"
# PASSWORD = "Indore@123"
executable_path = os.getcwd()+"\\geckodriver_v0.30.exe"
# logging 
today_date=date.today()
logfile =  os.getcwd() +'\\logs\\'+'coderbyte_logfile'+str(today_date)+'.txt'
logging.basicConfig(level=logging.INFO,filename=logfile,filemode='w',format='[line :- %(lineno)d] %(asctime)s [%(levelname)s] - %(message)s ')

locations_list=[]
locations_list.append(logfile)
profile = webdriver.FirefoxProfile()
# profile.set_preference("browser.helperApps.neverAsk.saveToDisk","text/plain,application/pdf")
# profile.set_preference("browser.download.manager.showWhenStarting",False)
# profile.set_preference('browser.download.dir',os.getcwd() + '\\'+'temp')
# profile.set_preference("browser.download.folderList",2)
# profile.set_preference('pdfjs.disabled',True)
# mime_types = ['text/plain', 'application/vnd.ms-excel', 'text/csv', 'application/csv', 'text/comma-separated-values',
# 				  'application/download', 'application/octet-stream', 'binary/octet-stream', 'application/binary', 'application/x-unknown']
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
# profile.set_preference("security.tls.version.min", 1)
profile.accept_untrusted_certs = True

# binary = FirefoxBinary(r"C:\\Program Files\\Mozilla Firefox\\Firefox.exe")
# binary = FirefoxBinary(r"C:\\Users\\chetan.surwade\\AppData\\Local\\Mozilla Firefox\\firefox.exe")
# driver = webdriver.Firefox(firefox_binary=binary, firefox_profile=profile,
#                             options=options, executable_path=executable_path)
driver = webdriver.Firefox(firefox_profile=profile,
                                        options=options, executable_path=executable_path)

# driver = webdriver.Firefox(executable_path=GeckoDriverManager().install(),firefox_profile=profile)
# driver = webdriver.Firefox(executable_path=executable_path,firefox_profile=profile, options=Options)

# get credentials from bu_config
credential_dict = get_config('CODER_BYTE_INVOICE_AUTOMATION','CODER_BYTE_INVOICE_AUTOMATION')
username = credential_dict['USERNAME'].split(';')[0]
password = credential_dict['PASSWORD'].split(';')[0]
sp_username = credential_dict['USERNAME'].split(';')[1]
sp_password =  credential_dict['PASSWORD'].split(';')[1]
share_point_path = credential_dict['API_KEY']
receiver_email = credential_dict['EMAIL_LIST']
# receiver_email = 'praveen.patel@biourja.com'
# receiver_email2= 'imam.khan@biourja.com'


job_name = 'CODER_BYTE_INVOICE_AUTOMATION'

def send_mail(receiver_email: str, mail_subject: str, mail_body: str, attachment_locations: list = None, sender_email: str = None, sender_password: str=None) -> bool:
    """The Function responsible to do all the mail sending logic.

    Args:
        sender_email (str): Email Id of the sender.
        sender_password (str): Password of the sender.
        receiver_email (str): Email Id of the receiver.
        mail_subject (str): Subject line of the email.
        mail_body (str): Message body of the Email.
        attachment_locations (list, optional): Absolute path of the attachment. Defaults to None.

    Returns:
        bool: [description]
    """
    logging.info("INTO THE SEND MAIL FUNCTION")
    done = False
    try:
        logging.info("GIVING CREDENTIALS FOR SENDING MAIL")
        if not sender_email or sender_password:
            sender_email = "biourjapowerdata@biourja.com"
            sender_password = r"Texas08642"
            # sender_email = r"virtual-out@biourja.com"
            # sender_password = "t?%;`p39&Pv[L<6Y^cz$z2bn"
        receivers = receiver_email.split(",")
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = "biourjapowerdata@biourja.com"
        msg['To'] = receiver_email
        msg['Subject'] = mail_subject
        body = mail_body
        logging.info("Attaching mail body")
        msg.attach(email.mime.text.MIMEText(body, 'html'))
        logging.info("Attching files in the mail")
        for files_locations in attachment_locations:
            with open(files_locations, 'r+b') as attachment:
                # instance of MIMEBase and named as p
                p = email.mime.base.MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                p.set_payload((attachment).read())
                encoders.encode_base64(p)  # encode into base64
                p.add_header('Content-Disposition',
                             "attachment; filename= %s" % files_locations)
                msg.attach(p)  # attach the instance 'p' to instance 'msg'

        # s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s = smtplib.SMTP('smtp.office365.com',
                         587)  # creates SMTP session
        s.starttls()  # start TLS for security
        s.login(sender_email, sender_password)  # Authentication
        text = msg.as_string()  # Converts the Multipart msg into a string

        s.sendmail(sender_email, receivers, text)  # sending the mail
        s.quit()  # terminating the session
        done = True
        logging.info("Email sent successfully")
        print("Email sent successfully.")
    except Exception as e:
        print(
            f"Could not send the email, error occured, More Details : {e}")
    finally:
        return done


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
        logging.info('clicking on next button')
        driver.find_element(By.CSS_SELECTOR,'.btn.btn-gradient.employer.nextButton').click()
        logging.info('pass password')
        driver.find_element(By.XPATH,'//*[@id="app"]/div[2]/div/div/div[1]/div[1]/div[2]/div/input').send_keys(password)
        logging.info('click on Login Button')
        driver.find_element(By.CSS_SELECTOR,"button[class='btn btn-gradient employer']").click() # login button
        time.sleep(8)
        
        try:
            logging.info('Open plan and billing tab')
            driver.get('https://coderbyte.com/dashboard/biourja-efzrr#settings-plan_and_billing')
            time.sleep(5)
            driver.refresh()
            time.sleep(5)
            # WebDriverWait(driver,200,poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/section[6]/div/div[3]/div[2]/div/ul/li[2]/div[1]/a/span')))
            
        except:
            try:
                # driver.refresh()
                driver.get('https://coderbyte.com/dashboard/biourja-efzrr#settings-plan_and_billing')
                # WebDriverWait(driver,200,poll_frequency=2).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/section[6]/div/div[3]/div[2]/div/ul/li[2]/div[1]/a/span')))
                # # time.sleep(120)
            except Exception as e:
                raise e

        if month == driver.find_element(By.XPATH,'/html[1]/body[1]/div[1]/section[7]/div[1]/div[3]/div[1]/div[2]/ul[1]/li[3]/div[1]/span[1]').text.split()[0]:
            driver.find_element(By.XPATH,'/html[1]/body[1]/div[1]/section[7]/div[1]/div[3]/div[1]/div[2]/ul[1]/li[3]/div[3]/a[1]').click()  #latest invoice
        else:
            logging.info("Sending mail for JOB FAILED")
            bu_alerts.send_mail(receiver_email = receiver_email,mail_subject ='JOB FAILED - {}'.format(job_name),mail_body = 'No new invoice found for {}'.format(month),attachment_location = logfile)
            sys.exit() 
        time.sleep(1)   
        driver.switch_to.window(driver.window_handles[-1])
        # WebDriverWait(driver,90).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div/div/div[1]/div/div[3]/div[1]/div/div[2]/table/tbody/tr[4]/td/div/button[1]/div/span')))
        # logging.info('Click on download button')
        # driver.find_element_by_xpath('/html/body/div/div/div[1]/div/div[3]/div[1]/div/div[2]/table/tbody/tr[4]/td/div/button[1]/div/span').click()  #download invoice
        time.sleep(1)
        logging.info('*****************Download Successfully*************')
        logging.info(f'File is downloaded in {download_wait()} seconds.')
        s = connect_to_sharepoint()
        file_upload_sp(s)
        logging.info("File is uploaded")
        logging.info("Sending mail for JOB SUCCESS")
        send_mail(receiver_email = receiver_email,mail_subject =f'JOB SUCCESS - {job_name}',mail_body = f'{job_name} completed successfully , Attached PDF and Logs \n {share_point_path}',attachment_locations = locations_list)
        logging.info("EMail Sent Successfully")
        
    except Exception as e:
        logging.exception(str(e))
        logging.info("Sending mail for JOB FAILED")
        bu_alerts.send_mail(receiver_email = receiver_email,mail_subject ='JOB FAILED - {}'.format(job_name),mail_body = 'Error in main() details {}'.format(str(e)),attachment_location = logfile)
    finally:
        driver.quit()

if __name__=="__main__" :
    logging.info('Execution Started')
    time_start = time.time()
    
    main()
    
    time_end = time.time()
    logging.info(f"It takes {time_end-time_start} seconds to run")
    logging.info("Execution ended")