import os
import time
import sys
import traceback
from datetime import datetime  
from datetime import timedelta  
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver.common.by import By  

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class ElementLocatorNotExists(Exception):
    def __init__(self, message):
        self.message = message
    def __str__(self):
        return self.message

class CourtNumberNotExists(Exception):
   """Raised when cannot find the file"""
   pass
 
class TimeNotAvailable(Exception):
   """Raised when cannot find the file"""
   pass
     
def check_exists_by_xpath(xpath):
    try:
        driver.find_element(By.XPATH, xpath)
        ## driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True
      
def check_exists_by_class_name(class_name):
    try:
        driver.find_element(By.CLASS_NAME, class_name)
        ## driver.find_element_by_class_name(class_name)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_name(name):
    try:
        driver.find_element(By.NAME, name)
        ## driver.find_element_by_name(name)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_id(id):
    try:
        driver.find_element(By.ID, id)
        ## driver.find_element_by_id(id)
    except NoSuchElementException:
        return False
    return True

def get_clock_hour(military_hour):
    if (military_hour>12):
        clock_hour="00"+str(military_hour-12)
        clock_hour=clock_hour[-2:]+":00 PM"
    elif (military_hour==12):
        clock_hour=str(military_hour)+":00 PM"
    else:
        clock_hour= "00"+str(military_hour)
        clock_hour=clock_hour[-2:]+":00 AM"
    return(clock_hour)

def log_process(file_path, message):
    ##If file exists, append it
    if os.path.exists(file_path):
        ## os.remove(file_path)        
        direct_writing_file=open(file_path, "a")
    else:
        direct_writing_file=open(file_path, "w")

    ##Write to the file
    direct_writing_file.write(message +"\n")
    direct_writing_file.close()

def sent_email(message):
    sender_email = 'waynelxb@gmail.com'
    sender_email_password='WAYdl@1125'
    receiver_email=sender_email
    server = 'smtp.gmail.com'
    port = 587
    msg = MIMEMultipart("alternative")
    msg["Subject"] = 'Book Tennis Court'
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.attach(MIMEText(message, 'plain'))    
    s = smtplib.SMTP(server, port)
    s.ehlo()
    s.starttls()
    s.login(sender_email, sender_email_password)
    s.sendmail(sender_email, receiver_email, msg.as_string())
    s.quit()        


 
################## Static Input Parameter: Account and Appoinment Info ##########################    
if len(sys.argv) == 2:  
    court_number=sys.argv[1]
    login_email_alias=""
    str_appointment_military_hour=""
elif len(sys.argv) == 3:   
    court_number=sys.argv[1]
    login_email_alias=sys.argv[2]
    str_appointment_military_hour=""
elif len(sys.argv) == 4: 
    court_number=sys.argv[1]
    login_email_alias=sys.argv[2]    
    str_appointment_military_hour=sys.argv[3]
else:
    raise IndexError("The length of sys.argv should not be less than 2 or more than 4.")    
    
## login_email_alias="w"
## court_number="3"
## str_appointment_military_hour = "[1,2]"


######## get the appointment date
now =  datetime.now()
now_string=now.strftime("%Y%m%d%H%M%S")
appointment_date=(now + timedelta(days=7)).date()
str_appointment_day_of_month=str(appointment_date.day)


######## login email 
    
if login_email_alias!="":     
    if login_email_alias=="w":
        login_email="waynelxb@gmail.com"
        user_name="Eric"
    if login_email_alias=="b":
        login_email="befwaynelxb@gmail.com"
        user_name="befwaynelxb"    
    if login_email_alias=="u":
        login_email="liuxinbo.utube@gmail.com"
        user_name="Utube"
    if login_email_alias=="a":
        login_email="apulxb@gmail.com"
        user_name="Apu"
else:
    if appointment_date.weekday()==0:
        login_email="waynelxb@gmail.com"
    elif appointment_date.weekday()==1:
        login_email="apulxb@gmail.com"
    elif appointment_date.weekday()==2:
        login_email="befwaynelxb@gmail.com"        
    elif appointment_date.weekday()==3:
        login_email="befwaynelxb@gmail.com"       
    elif appointment_date.weekday()>=5:
        login_email="liuxinbo.utube@gmail.com"    
    
    
    
######## one password for all login emails     
login_password="TENdl@1125"

######## convert court_number to court_name
if (court_number=="2"):
    court_name="Court 2 (Center Court)"
elif (court_number=="3"):
    court_name="Court 3 (North Court)"
else:
    raise CourtNumberNotExists()



####### get the hour list
n=len(str_appointment_military_hour)
if(n>2):
    ######## if n>2 then use the input hour list    
    list_appointment_military_hour=str_appointment_military_hour[1:n-1].split(',')
else:
    ####### use the default hour list    
    if(appointment_date.weekday()>=5):
        list_appointment_military_hour=[10,11,15,16,17,18,19]
    else:
        list_appointment_military_hour=[16,17,18,19]




####### get the hour list
n=len(str_appointment_military_hour)
if(n>2):
    ######## if n>2 then use the input hour list    
    list_appointment_military_hour=str_appointment_military_hour[1:n-1].split(',')
else:
    ####### use the default hour list    
    if(appointment_date.weekday()>=5):
        list_appointment_military_hour=[10,11,15,16,17,18,19]
    else:
        list_appointment_military_hour=[16,17,18,19]



print(login_email)
print(court_name)
print(list_appointment_military_hour)   


#################### Static Input Parameter: Program Configuration ##########################
log_path="C:/Projects/Python/BookTennisCourt/BookTennisCourtLog.txt"    
chrome_driver_path = "C:\Projects\Python/chromedriver.exe"   
profile_directory=r"–user-data-dir=C:\Users\Me\AppData\Local\Google\Chrome\User Data"
login_url="https://booking.appointy.com/en-US/tennisonthelake/bookings/login"
implicitly_wait_second=5
sleep_second=3

######## generate summary message
msg_summary="\n\n\n"+"Book Time: "+now.strftime("%Y-%m-%d %H:%M:%S") + "\n"
msg_summary=msg_summary+"Login Email: "+login_email+"\n"
msg_summary=msg_summary+"Court Name: " + court_name +"\n"
msg_summary=msg_summary+"Expected Date: "+str(appointment_date)+" "+appointment_date.strftime('%A')+"\n"
msg_summary=msg_summary+"Expected Hour List: "+str(list_appointment_military_hour)+"\n"

    
try:     
    now =  datetime.now()   
   
    svc=Service(chrome_driver_path)
    options = webdriver.ChromeOptions()
    options.add_argument(profile_directory)
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_experimental_option("excludeSwitches", ["enable-logging"])       
    driver = webdriver.Chrome(service=svc, options=options)
    driver.implicitly_wait(implicitly_wait_second) ## seconds  
    
######## Initialize driver with login_url to work on login page 
    driver.get(login_url); 
    ## element_for_checking_email = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "email")))
    id_element_email_inputbox="email"    
    if check_exists_by_id(id_element_email_inputbox) == True:
        element_email_inputbox=driver.find_element(By.ID, id_element_email_inputbox)
        element_email_inputbox.send_keys(login_email)
    else:
        raise ElementLocatorNotExists(id_element_email_inputbox)
    time.sleep(sleep_second)
    
######## Click the button "NEXT" on login page    
    #### <label class="btn btn-primary ng-tns-c1-1 ng-star-inserted" tabindex="0" for="submit-formRegister" style=""><!----><!----><!----><span class="ng-tns-c1-1">Next</span><!---->  
    #### Method 1
    ## driver.find_element_by_xpath("//label[@class='btn btn-primary ng-tns-c1-1 ng-star-inserted']/span[text()='Next']").click()
    #### Method 2    
    xpath_element_register_button="//label[@for='submit-formRegister']/span[text()='Next']"    
    if check_exists_by_xpath(xpath_element_register_button) == True:
        element_register_button=driver.find_element(By.XPATH, xpath_element_register_button)
        element_register_button.click()        
        msg_email_register="Login email has been registered.\n"         
        msg_summary=msg_summary+msg_email_register
    else:
        raise ElementLocatorNotExists(xpath_element_register_button)
    time.sleep(sleep_second)

######## Switch to the authentication page
    driver.switch_to.window(driver.window_handles[1])
    #### Input password 
    class_name_element_password_inputbox="password-input"
    if check_exists_by_class_name(class_name_element_password_inputbox) == True:
        element_password_inputbox=driver.find_element(By.CLASS_NAME, class_name_element_password_inputbox)
        element_password_inputbox.send_keys(login_password)
    else:
        raise ElementLocatorNotExists(class_name_element_password_inputbox)
        
    time.sleep(sleep_second)   
    #### Click the button CONTINUE
    xpath_element_password_submit_button='//button[@type="submit"]'
    element_password_submit_button=driver.find_element(By.XPATH, xpath_element_password_submit_button)        
    element_password_submit_button.click()
    msg_send_password="Login password has been entered.\n"
    msg_summary=msg_summary+msg_send_password
    time.sleep(sleep_second)

######## Switch to the court page to pick a court
    driver.switch_to.window(driver.window_handles[-1])
    ## print(driver.current_url)
    time.sleep(sleep_second) 
    
    #### Pick court   
    xpath_element_court_name="//span[text()='"+court_name+"']"
    ## element_court_option=driver.find_element_by_xpath("//div[@class='ng-tns-c4-3']/span[text()='Court 3 (North Court)']") 
    element_court_option=driver.find_element(By.XPATH, xpath_element_court_name)
    element_court_option.click()
    time.sleep(sleep_second)    
    
    #### Click NEXT button to confirm court selection        
    xpath_element_court_confirmation_button="//button[@type='button'][@id='app-next-button']"
    element_count_confirmation_button=driver.find_element(By.XPATH, xpath_element_court_confirmation_button)
    element_count_confirmation_button.click()
    msg_pick_court=court_name + " has been confirmed.\n"
    msg_summary=msg_summary+msg_pick_court
    
######## Switch to calendar page to pick date and time                  
    driver.switch_to.window(driver.window_handles[-1])     
    ## print(driver.current_url)     
    time.sleep(sleep_second)     
    #### Pick date
    # print(str_appointment_day_of_month)
    xpath_element_date_option="//button[contains(@class, 'btn-link ng-tns-')][@tabindex='0']/span[text()='"+str_appointment_day_of_month+"']"
    if check_exists_by_xpath(xpath_element_date_option) == True:
        element_date_option=driver.find_element(By.XPATH, xpath_element_date_option)
        element_date_option.click()
        msg_pick_date=str(appointment_date) + " for "+court_name+" has been selected.\n"
        msg_summary=msg_summary+msg_pick_date   
    else:
        msg_pick_date=str(appointment_date) +" is not available for "+ court_name +", or the element cannot be found.\n"
        msg_summary=msg_summary+msg_pick_date   
        ## print(message)             
        ## log_process(log_path, msg_pick_date)
        raise ElementLocatorNotExists(xpath_element_date_option)      
    time.sleep(sleep_second) 
    
    #### Pick time slot    
    is_time_slot_available=False    
    for appointment_military_hour in  list_appointment_military_hour:
        appointment_clock_hour=get_clock_hour(int(appointment_military_hour))        
        xpath_element_time_option="//button[contains(@class, 'btn-link app-calendar-time-item available')][@tabindex='0']/span[text()='"+appointment_clock_hour+"']"    
        if check_exists_by_xpath(xpath_element_time_option) == True:
            element_time_option=driver.find_element(By.XPATH, xpath_element_time_option)
            element_time_option.click()          
            msg_pick_time=str(appointment_date) + " "+ appointment_clock_hour + " is available and has been selected for "+ court_name+"\n"
            msg_summary=msg_summary+msg_pick_time  
            ## print(message)
            is_time_slot_available=True
            break
        else: 
            msg_pick_time = str(appointment_date) + " "+ appointment_clock_hour +  " is NOT available for "+ court_name+"\n"
            msg_summary=msg_summary+msg_pick_time  
            ## log_process(log_path, message)
            ## print(message)        
    if is_time_slot_available==False:
        raise TimeNotAvailable()   
    time.sleep(sleep_second)
      
######## Switch to user name input page
    driver.switch_to.window(driver.window_handles[-1])    
    ## print(4)          
    ## print(driver.current_url)     
    #### Input user name
    xpath_element_user_name="//input[@type='text']"
    element_input_user_name=driver.find_element(By.XPATH, xpath_element_user_name)    
    element_input_user_name.send_keys(user_name)
   
    #### Click the button NEXT to confirm user name   
    xpath_element_user_name_confirmation_button="//label[@for='submit-formIntake']/span[text()='Next']"
    element_user_name_confirmation_button=driver.find_element(By.XPATH, xpath_element_user_name_confirmation_button)
    element_user_name_confirmation_button.click()
    
    msg_confirm_user_name="User name " + user_name +" has been confirmed.\n"
    msg_summary=msg_summary+msg_confirm_user_name  
    ## log_process(log_path, "User name " + user_name +" has been confirmed.")
    time.sleep(sleep_second)

######## Switch to final confirmation page
    driver.switch_to.window(driver.window_handles[-1])   
    ## print(driver.current_url) 
    #### Click CONFIRM button
    xpath_element_final_confirmation_button="//button[@id='app-next-button']"
    element_final_confirmation_button=driver.find_element(By.XPATH, xpath_element_final_confirmation_button)
    element_final_confirmation_button.click()    
    time.sleep(sleep_second)    
    
######## Switch to Confirmed Appoiment Details page
    driver.switch_to.window(driver.window_handles[-1])   
    ## print(driver.current_url) 
    xpath_element_startover_button="//span[text()='Start Over']"   
    if check_exists_by_xpath(xpath_element_startover_button) == True:    
        #### No need to click StartOver, the appearance of this page means the appointment has been booked.
        msg_confirm_booking = str(appointment_date)+" "+appointment_clock_hour+" on "+court_name+" has been booked successfully.\n"    
        msg_summary=msg_summary+msg_confirm_booking  
        log_process(log_path, msg_summary)
        driver.quit() 
        sent_email(msg_summary)
    else:
        raise ElementLocatorNotExists(xpath_element_startover_button)

except ElementLocatorNotExists as e:
    msg_summary=msg_summary+"Error: "+ e.message +" cannot be loacated.\n"
    print(msg_summary)
    driver.quit()
    log_process(log_path, msg_summary)    
    sent_email(msg_summary)
    
except CourtNumberNotExists:
    message=msg_summary+ "Error: "+ court_number + " doesn't exist.\n"     
    print(message)
    log_process(log_path, message)
    sent_email(msg_summary)

   
except TimeNotAvailable:
    msg_summary=msg_summary+"Error: No time slot in the list "+str(list_appointment_military_hour)+" is available.\n" 
    driver.quit()    
    print(msg_summary)   
    log_process(log_path, msg_summary)
    sent_email(msg_summary)
  
except: 
    exc_type, exc_value, exc_traceback = sys.exc_info()
    exceptMessage=repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
    msg_summary=msg_summary+"Error: "+exceptMessage +"\n"
    driver.quit()
    print(msg_summary)   
    log_process(log_path, msg_summary)
    sent_email(msg_summary)

