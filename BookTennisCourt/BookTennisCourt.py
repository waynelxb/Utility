###########################
##### Version 2.0 #########
###########################

import os
import time
import sys
import traceback
from datetime import datetime  
from datetime import timedelta  
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
# from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver.common.by import By  
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sqlite3
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class CourtOverbooked(Exception):
   """Raised when cannot find the file"""
   pass


class EmailNotUsable(Exception):
   """Raised when cannot find the file"""
   pass

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

def get_element_wait_for_load(element_type,element_expression):
    wait=WebDriverWait(driver, 10)  
    # element = driver.find_element(By.XPATH, element_expression);
    try:     
        if element_type=="XPATH":
            element=wait.until(EC.presence_of_element_located((By.XPATH, element_expression)))
        if element_type=="ID":
            element=wait.until(EC.presence_of_element_located((By.ID, element_expression)))
        if element_type=="CLASS_NAME":
            element=wait.until(EC.presence_of_element_located((By.CLASS_NAME, element_expression)))
        if element_type=="NAME":
            element=wait.until(EC.presence_of_element_located((By.NAME, element_expression)))
        return element
    except:
        return "None"

def check_element_existence(element_type,element_expression):
    try:    
        if element_type=="XPATH":
            driver.find_element(By.XPATH, element_expression)
        if element_type=="ID":
            driver.find_element(By.ID, element_expression)
        if element_type=="CLASS_NAME":
            driver.find_element(By.CLASS_NAME, element_expression)
        if element_type=="NAME":
            driver.find_element(By.NAME, element_expression)
        return True
    except:
        return False      
        
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


def send_email(message, status):
    sender_email = 'waynelxb@gmail.com'
    sender_email_password='WAYdl@1125'
    receiver_email=sender_email
    server = 'smtp.gmail.com'
    port = 587
    msg = MIMEMultipart("alternative")
    msg["Subject"] = status+ ' in Booking Tennis Court'
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.attach(MIMEText(message, 'plain'))    
    s = smtplib.SMTP(server, port)
    s.ehlo()
    s.starttls()
    s.login(sender_email, sender_email_password)
    s.sendmail(sender_email, receiver_email, msg.as_string())
    s.quit()   
    

############### DB Operation Functions ################################################

def create_tennis_court_booking_db_connection(db_file):
    conn = sqlite3.connect(db_file)
    return conn

def create_appointment_table(conn, script):
    cur=conn.cursor()
    cur.execute(script)

def sqlite_purge_appointment(conn):    
    cur=conn.cursor()  
    purge_record_script="delete from Appointment"
    cur.execute(purge_record_script);   
    conn.commit()                      
    cur.close()

def sqlite_insert_appointment(conn, batch_id, login_email, login_time, appt_time, court_number, appt_status):
    insert_query="INSERT INTO Appointment (BatchID, LoginEmail, LoginTime, AppointmentTime, CourtNumber, AppointmentStatus)"+\
                 " VALUES("+str(batch_id)+ ",'"+login_email+"','"+login_time+ "','"+appt_time+"',"+str(court_number)+",'"+appt_status+"')"                    
    # print(insert_query)             
    conn.execute(insert_query)
    conn.commit()
   

def sqlite_get_appointment(conn):    
    cur=conn.cursor()  
    ### get all appointment week records    
    select_all_query="select * from Appointment"
    cur.execute(select_all_query)  
    query_result=cur.fetchall()                                  
    cur.close()
    return("Appointment Records:\n"+ str(query_result).replace("),", ")\n") ).replace("[","").replace("]","")


def sqlite_delete_old_appointment(conn, appt_time):    
    cur=conn.cursor()     
    dt_appt_time = datetime.strptime(appt_time, '%Y-%m-%d') 
    appt_week_start_date = str(dt_appt_time-timedelta(days=dt_appt_time.weekday()))  # Monday is the start date of a week        
    ### delete old records which are not in the curent appointment week
    delete_query="delete from Appointment "+\
                  "where AppointmentTime <'" + appt_week_start_date+"'"
    # print(delete_query)
    cur.execute(delete_query);   
    conn.commit()           
    cur.close()


def sqlite_check_court_availability(conn, court_number, appt_time):    
    cur=conn.cursor()     
    #### if for current batch, there are two login emails not able to find an available time slot, then this court is not available
    query_check_court_availablity=" select CourtNumber " \
                                 +" from Appointment " \
                                 +" where AppointmentStatus='CourtOverbooked' and CourtNumber="+str(court_number) + " and AppointmentTime='"+appt_time+"'"                                 
    cur.execute(query_check_court_availablity)      
    query_result=cur.fetchone()  
    cur.close()        
    if str(query_result)=="None":        
        return(True)
    else:
        return(False)


def sqlite_check_email_usability(conn, email, appt_time, court_number):    
    cur=conn.cursor()     
    dt_appt_time = datetime.strptime(appt_time, '%Y-%m-%d') 
    appt_week_start_date = str(dt_appt_time-timedelta(days=dt_appt_time.weekday()))  # Monday is the start date of a week
    #### check whether the email has been used more than 3 times in the booking week, it doesn't matter on which date or court.   
    check_email_query=" select LoginEmail " \
                     +" from Appointment  " \
                     +" where AppointmentStatus='EmailOverused' and LoginEmail='"+email + "'" \
                     +" union "  \
                     +" select LoginEmail " \
                     +" from Appointment " \
                     +" where AppointmentStatus='Succeeded' and LoginEmail='"+email + "' and AppointmentTime>='" + appt_week_start_date +"'" \
                     +" group by LoginEmail having count(1)>=3 "
    # check_email_query=check_email_query              
    cur.execute(check_email_query) 
    query_result=cur.fetchone()      
    cur.close()    
    # print(check_email_query)
    # print("check_email_query result: "+ str(query_result))
    if str(query_result)=="None":
        return(True)
    else:
        return(False)   


#################################################################################################
################## Static Input Parameter: Account and Appoinment Info ##########################
#################################################################################################
if len(sys.argv) == 2:  
    court_number=sys.argv[1]
    str_military_hour_option=""
elif len(sys.argv) == 3:   
    court_number=sys.argv[1]
    str_military_hour_option=sys.argv[2] 
else:
    raise IndexError("The length of sys.argv should not be less than 2 or more than 3.")    


#########>>>>>>>>>>>>> Testing <<<<<<<<<<<<<<<#######
court_number="3"
str_military_hour_option=""    
# # str_military_hour_option = "[5,6]"    

##########################################################################################################
###### If enable_purge_record =True, Appointment table will be purged. This variable is set manually #####
##########################################################################################################
enable_purge_record=False


###### login email list
list_email=["waynelxb@gmail.com","befwaynelxb@gmail.com","liuxinbo.utube@gmail.com","apulxb@gmail.com"]
# list_email=["waynelxb@gmail.com","liuxinbo.utube@gmail.com"]
######## One password for all login emails     
login_password="TENdl@1125"

######## Convert court_number to court_name 
if (court_number=="2"):
    court_name="Court 2 (Center Court)"
elif (court_number=="3"):
    court_name="Court 3 (North Court)"
else:
    raise CourtNumberNotExists() 

######## Generate batch id and msg_summary
batch_id=int(datetime.now().strftime("%Y%m%d%H%M%S"))
msg_summary="\n\n"+"BatchID: "+str(batch_id) + "\n"   

      
#################### Static Input Parameter: Program Configuration ##########################
##### Log file and db file are put in Data folder
log_path="C:/Projects/GitHub/BookTennisCourtData/BookTennisCourtLog.txt"    
db_file_path="C:/Projects/GitHub/BookTennisCourtData/TennisCourtBooking.db"

##### Chrome driver is put in project code folder
chrome_driver_path = "C:/Projects/GitHub/UtilityRepository/BookTennisCourt/chromedriver.exe"   

profile_directory=r"–user-data-dir=C:\Users\Me\AppData\Local\Google\Chrome\User Data"
login_url="https://booking.appointy.com/en-US/tennisonthelake/bookings/login"

create_table_script="""CREATE TABLE IF NOT EXISTS Appointment
                       (
                          BatchID              BIGINT,
                          LoginEmail           VARCHAR(50) NOT NULL,
                          LoginTime            TEXT,
                          AppointmentTime      TEXT,
                          CourtNumber          INT,
                          AppointmentStatus    VARCHAR(50)
                       )
                 """
# implicitly_wait_second=3
sleep_second=1

try:   
    ######## create a sqlite db and table to track the usage of login_email    
    conn = create_tennis_court_booking_db_connection(db_file_path)
    create_appointment_table(conn, create_table_script)   
    
    ######## Purge Appointment table           
    if enable_purge_record==True:        
        sqlite_purge_appointment(conn)

    ######## Test record 
    ### sqlite_insert_appointment(conn, 111111, 'waynelxb@gmail.com', '2022-2-14', '2022-2-21', 3, 'EmailOverused') 

    ######## At the end of the while loop, if is_email_usable=False, another login_email will be used.    
    is_email_usable=False
    while is_email_usable==False:   
        ##### Get login time        
        now=datetime.now()
        str_login_time=now.strftime("%Y-%m-%d %H:%M:%S") 
        appointment_date=(now + timedelta(days=7)).date()
        str_appointment_date=str(appointment_date)
        str_appointment_day_of_month=str(appointment_date.day)             
        
        ####### If the length of str_military_hour_option is longer than 2, then conver it to a list, or use the default hour list   
        input_option_length=len(str_military_hour_option)        
        if(input_option_length>2):             
            list_military_hour_option=str_military_hour_option[1:input_option_length-1].split(',')
        else:
            #### use the default hour list    
            if(appointment_date.weekday()>=5):
                list_military_hour_option=[9,10,11,12,14,15,16,17,18]
            else:
                list_military_hour_option=[17,16,18,15,14,13]   
                
        #### Delete the old records not in the current appoitment week  
        sqlite_delete_old_appointment(conn, str_appointment_date)

        #### Check whehter the court has been logged as overused  
        if sqlite_check_court_availability(conn, court_number, str_appointment_date)==False:
            raise CourtOverbooked()               
            
        #### Find usable login email   
        login_email=""              
        for email in list_email:
            # print(email)
            if sqlite_check_email_usability(conn, email, str_appointment_date, court_number)  == True:
                login_email=email  
                if login_email=="waynelxb@gmail.com":
                    user_name="Eric"
                if login_email=="befwaynelxb@gmail.com":
                    user_name="befwaynelxb"    
                if login_email=="liuxinbo.utube@gmail.com":
                    user_name="Utube"
                if login_email=="apulxb@gmail.com":
                    user_name="Apu" 
                is_email_usable=True
                break   
        #### all emails are overused
        if is_email_usable==False:
            raise EmailNotUsable()                  
    
        
        msg_summary=msg_summary+"Login Time: "+str_login_time + "\n"   
        msg_summary=msg_summary+"Login Email: "+login_email+"\n"
        msg_summary=msg_summary+"Court Name: " + court_name +"\n"
        msg_summary=msg_summary+"Appointment Date: "+str_appointment_date+" "+appointment_date.strftime('%A')+"\n"
        msg_summary=msg_summary+"Expected Hour List: "+str(list_military_hour_option)+"\n"


        ####### Create chrome driver       
        svc=Service(chrome_driver_path)
        options = webdriver.ChromeOptions()
        options.add_argument(profile_directory)
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_experimental_option("excludeSwitches", ["enable-logging"])       
        driver = webdriver.Chrome(service=svc, options=options)
        # driver.implicitly_wait(implicitly_wait_second) ## seconds  
        
        ######## Open email register page
        driver.get(login_url)
        id_element_email_inputbox="email" 
        element_email_inputbox=get_element_wait_for_load("ID",id_element_email_inputbox)      
        element_email_inputbox.send_keys(login_email)        
        
        ######## Confirm login email    
        xpath_element_register_button="//label[@for='submit-formRegister']/span[text()='Next']"         
        element_register_button=get_element_wait_for_load("XPATH",xpath_element_register_button)      
        element_register_button.click()        
        msg_email_register="Login email has been registered.\n"         
        msg_summary=msg_summary+msg_email_register        
    
        ######## Swith to the password page
        driver.switch_to.window(driver.window_handles[1])
        #### Input password 
        class_name_element_password_inputbox="password-input"        
        element_password_inputbox=get_element_wait_for_load("CLASS_NAME",class_name_element_password_inputbox)         
        element_password_inputbox.send_keys(login_password)        
        msg_email_register="Password has been input.\n"         
        msg_summary=msg_summary+msg_email_register            
        time.sleep(sleep_second)  
      
        #### Click the button to confirm password
        xpath_element_password_submit_button='//button[@type="submit"]'
        element_password_submit_button=get_element_wait_for_load("XPATH",xpath_element_password_submit_button)                 
        element_password_submit_button.click()               
    
        ######## Switch to the court page        
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))        
        driver.switch_to.window(driver.window_handles[-1])
        
        ######## Pick court   
        xpath_element_court_name="//span[text()='"+court_name+"']" 
        element_court_option=get_element_wait_for_load("XPATH",xpath_element_court_name)          
        element_court_option.click()   

        
        ######## Click NEXT button to confirm court selection        
        time.sleep(sleep_second)  
        xpath_element_court_confirmation_button="//button[@id='app-next-button']"
        element_count_confirmation_button=get_element_wait_for_load("XPATH",xpath_element_court_confirmation_button)            
        element_count_confirmation_button.click()
        msg_pick_court=court_name + " has been confirmed.\n"
        msg_summary=msg_summary+msg_pick_court             
        
        ######## Switch to calendar page to pick date and time  
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))                   
        driver.switch_to.window(driver.window_handles[-1])     
        
        #### make sure appointment_day_of_month show up on the page, then continue    
        xpath_element_appointment_day_of_month="//button[contains(@class, 'btn-link ng-tns-')]/span[text()='"+str_appointment_day_of_month+"']"      
        
        if get_element_wait_for_load("XPATH", xpath_element_appointment_day_of_month)!="None":          
            # time.sleep(sleep_second)             
            ### element for active date      
            ### <button class="btn-link ng-tns-c5-4 ng-star-inserted" position="bottom" data-tippy="" tabindex="0"><span class="ng-tns-c5-4">20</span><span class="app-calendar-month-name active"></button>
            xpath_element_active_date="//button[contains(@class, 'btn-link ng-tns-')][@tabindex='0']/span[text()='"+str_appointment_day_of_month+"']"
            ### element for inactive date due to overused email            
            ### <button class="btn-link ng-tns-c2-6 disabled ng-star-inserted" position="bottom" data-tippy="" tabindex="-1" data-original-title="You are allowed to book 3 service weekly "><span class="ng-tns-c2-6">18</span><span class="app-calendar-month-name active"><!----><span class="icon icon-calendar-insert ng-tns-c2-6 ng-star-inserted"></span><!----></span></button>                          
            xpath_element_no_date_available_due_to_login_email_overused="//button[contains(@class, 'btn-link ng-tns-')][@tabindex='-1'][@data-original-title='You are allowed to book 3 service weekly ']/span[text()='"+str_appointment_day_of_month+"']"     
           
            ### element for inactive date due to overbookd court
            ### <button class="btn-link ng-tns-c5-4 disabled ng-star-inserted" position="bottom" data-tippy="" tabindex="-1" data-original-title="Not Available"><span class="ng-tns-c5-4">20</span><span class="app-calendar-month-name active"><!----> Mar <!----></span></button>
            xpath_element_no_date_available_due_to_court_overbooked="//button[contains(@class, 'btn-link ng-tns-')][@tabindex='-1'][@data-original-title='Not Available']/span[text()='"+str_appointment_day_of_month+"']"    
      
            element_active_date=get_element_wait_for_load("XPATH", xpath_element_active_date)         
            # element_no_date_available_due_to_login_email_overused=get_element_wait_for_load("XPATH", xpath_element_no_date_available_due_to_login_email_overused)  
            # element_no_date_available_due_to_court_overbooked=get_element_wait_for_load("XPATH", xpath_element_no_date_available_due_to_court_overbooked)  
            
            if element_active_date != "None":
                # element_date_option=driver.find_element(By.XPATH, xpath_element_active_date)
                element_active_date.click()
                msg_pick_date=str(appointment_date) + " for "+court_name+" has been selected.\n"
                msg_summary=msg_summary+msg_pick_date  
                ### check whether the email has been overused, if yes, then try another login email             
            elif get_element_wait_for_load("XPATH", xpath_element_no_date_available_due_to_login_email_overused) != "None":
                driver.quit()                              
                sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_appointment_date, court_number, "EmailOverused")
                msg_pick_date="Description: " + login_email +" has been overused. Another login email will be used.\n"         
                msg_summary=msg_summary+msg_pick_date+sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                log_process(log_path, msg_summary)
                send_email(msg_summary, "Failed")            
                is_email_usable=False               
                ### then check whether the court has been overbooked       
            elif get_element_wait_for_load("XPATH", xpath_element_no_date_available_due_to_court_overbooked) != "None":    
                driver.quit()                                 
                sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_appointment_date, court_number, "CourtOverbooked")  
                raise CourtOverbooked()
            else:
                raise ElementLocatorNotExists("The element xpath of appointment day of month " + str_appointment_day_of_month)             
            
    ############### This part need to enhance speed use time item  
    # <button class="btn-link app-calendar-time-item available" container="body" placement="left" triggers="manual" data-tippy="" data-original-title="2022-02-16T05:32:41.908Z" tabindex="0"><span class="app-calendar-time">05:00 AM</span><!----></button>
    xpath_element_app_calendar_time_item="//button[contains(@class, 'app-calendar-time-item')]"      
    if get_element_wait_for_load("XPATH", xpath_element_app_calendar_time_item)!="None":        
        #### Pick time slot    
        is_time_slot_available=False    
        for appt_military_hour in list_military_hour_option:
            appt_clock_hour=get_clock_hour(int(appt_military_hour))        
            xpath_element_time_option="//button[contains(@class, 'btn-link app-calendar-time-item available')][@tabindex='0']/span[text()='"+appt_clock_hour+"']"   
            # element_time_option=get_element_wait_for_load("XPATH", xpath_element_time_option)              
            if check_element_existence("XPATH", xpath_element_time_option) == True:
                element_time_option=driver.find_element(By.XPATH, xpath_element_time_option)
                element_time_option.click()    
                str_appointment_time=str_appointment_date + " "+ appt_clock_hour
                msg_pick_time=str_appointment_time + " is available and has been selected for "+ court_name+"\n"
                msg_summary=msg_summary+msg_pick_time  
                is_time_slot_available=True
                break
            else: 
                msg_pick_time = str(appointment_date) + " "+ appt_clock_hour +  " is NOT available for "+ court_name+"\n"
                msg_summary=msg_summary+msg_pick_time         
        if is_time_slot_available==False:
            raise TimeNotAvailable()   
    else:
      raise ElementLocatorNotExists(xpath_element_app_calendar_time_item)     

      
######## Switch to the page to confirm user name
    driver.switch_to.window(driver.window_handles[-1])   
    #### Input user name
    xpath_element_user_name="//input[@type='text']"  
    element_input_user_name=get_element_wait_for_load("XPATH",xpath_element_user_name)  
    element_input_user_name.send_keys(user_name)
   
    #### Click the button NEXT to confirm user name   
    xpath_element_user_name_confirmation_button="//label[@for='submit-formIntake']/span[text()='Next']"
    element_user_name_confirmation_button=get_element_wait_for_load("XPATH",xpath_element_user_name_confirmation_button)      
    # element_user_name_confirmation_button=driver.find_element(By.XPATH, xpath_element_user_name_confirmation_button)
    element_user_name_confirmation_button.click()    
    msg_confirm_user_name="User name " + user_name +" has been confirmed.\n"
    msg_summary=msg_summary+msg_confirm_user_name  

######## Switch to final confirmation page
    driver.switch_to.window(driver.window_handles[-1])   
    #### Click CONFIRM button
    xpath_element_final_confirmation_button="//button[@id='app-next-button']"
    # element_final_confirmation_button=driver.find_element(By.XPATH, xpath_element_final_confirmation_button)
    element_final_confirmation_button=get_element_wait_for_load("XPATH",xpath_element_final_confirmation_button)      
    element_final_confirmation_button.click()    
    
######## Switch to the appoitment receipt(StartOver) page
    driver.switch_to.window(driver.window_handles[-1])   
    xpath_element_startover_button="//span[text()='Start Over']"     
    # wait_for_load("XPATH", xpath_element_startover_button)      
    if get_element_wait_for_load("XPATH", xpath_element_startover_button) != "None":    
        #### No need to click StartOver, the appearance of this page means the appointment has been booked.
        msg_confirm_booking = str_appointment_time +" on "+court_name+" has been booked successfully.\n"
        sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_appointment_time, court_number,"Succeeded")       
        msg_summary=msg_summary+msg_confirm_booking+sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_process(log_path, msg_summary)
        driver.quit()     
        send_email(msg_summary, "Succeeded")
    else:
        raise ElementLocatorNotExists(xpath_element_startover_button)


except CourtOverbooked:
    msg_summary=msg_summary+"Exception: " +court_name+" has been overbooked.\n"+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"+ sqlite_get_appointment(conn)
    print(msg_summary)
    log_process(log_path, msg_summary)    
    send_email(msg_summary, "Failed")   

except EmailNotUsable:
    msg_summary=msg_summary+"Exception: All the emails in "+str(list_email)+" have been overused!\n"+sqlite_get_appointment(conn)
    print(msg_summary)
    log_process(log_path, msg_summary)    
    send_email(msg_summary, "Failed")    
    
except ElementLocatorNotExists as e:
    driver.quit()
    msg_summary=msg_summary+"Error: "+ e.message +" cannot be loacated.\n" +"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+sqlite_get_appointment(conn)
    print(msg_summary)
    log_process(log_path, msg_summary)    
    send_email(msg_summary, "Failed")
    
except CourtNumberNotExists:
    message=msg_summary+"Error: Court "+ court_number + " doesn't exist.\n"+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+sqlite_get_appointment(conn) 
    print(message)
    log_process(log_path, message)
    send_email(msg_summary, "Failed")
    
except TimeNotAvailable:
    driver.quit()
    message=msg_summary+"Error: No time slot in "+ str(list_military_hour_option) + " is available.\n"+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+sqlite_get_appointment(conn) 
    print(message)
    log_process(log_path, message)
    send_email(msg_summary, "Failed")   

except: 
    driver.quit()    
    exc_type, exc_value, exc_traceback = sys.exc_info()
    exceptMessage=repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
    msg_summary=msg_summary+"Error: "+exceptMessage +"\n" +"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+sqlite_get_appointment(conn)
    print(msg_summary)   
    log_process(log_path, msg_summary)
    send_email(msg_summary, "Failed")


