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
from CreateCalendarEvent.cal_setup import get_calendar_service
import socket



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



def create_calendar_event(event_date, event_hour, event_summary):
    try:
       # creates one hour event tomorrow 10 AM IST
       service = get_calendar_service()
       event_datetime = datetime(event_date.year, event_date.month, event_date.day, event_hour)+timedelta(days=0)
       start = event_datetime.isoformat()
       end = (event_datetime + timedelta(hours=1)).isoformat()
    
       event_result = service.events().insert(calendarId='primary'
          ,body={
                   'summary': event_summary,
                   'description': 'Tennis Court',
                   'start': {"dateTime": start, 'timeZone': 'America/Chicago'},
                   'end': {"dateTime": end, 'timeZone': 'America/Chicago'},         
                   'reminders': {'useDefault': False,
                                 'overrides': [{'method': 'email', 'minutes': 24 * 60},
                                               {'method': 'email', 'minutes': 8 * 60},
                                               {'method': 'popup', 'minutes': 120},
                                              ],
                                }
               }
          ,sendUpdates='all'
       )
       event_result.execute()
       return("Calender Event: Succeeded in creating event.\n")
       # print("created event")
       # print("id: ", event_result['id'])
       # print("summary: ", event_result['summary'])
       # print("starts at: ", event_result['start']['dateTime'])
       # print("ends at: ", event_result['end']['dateTime'])
    except:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        exceptMessage=repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
        return("Calender Event: "+exceptMessage +"\n")


def get_element_wait_for_load(wait_seconds,element_type,element_expression):
    wait=WebDriverWait(driver, wait_seconds)  
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


def send_email(subject, message):
    sender_email = 'waynelxb@gmail.com'
    sender_email_password='LXBdl@33#5708'
    receiver_email=sender_email
    server = 'smtp.gmail.com'
    port = 587
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
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

def sqlite_insert_appointment(conn, batch_id, login_email, login_time, appt_time, court_number, app_desc, appt_status):
    insert_query="INSERT INTO Appointment (BatchID, LoginEmail, LoginTime, AppointmentTime, CourtNumber, Description, AppointmentStatus)"+\
                 " VALUES("+str(batch_id)+ ",'"+login_email+"','"+login_time+ "','"+appt_time+"',"+str(court_number)+",'"+app_desc+"','"+appt_status+"')"                    
    # print(insert_query)             
    conn.execute(insert_query)
    conn.commit()

def sqlite_get_appointment(conn):    
    cur=conn.cursor()  
    ### get all appointment week records    
    select_all_query="select LoginEmail, strftime('%m/%d %H:00', AppointmentTime)||REPLACE(REPLACE(REPLACE(Description,'WeekDay:',''),' | ','-'),'Code: ','')||'-C'||CourtNumber AS Brief,  AppointmentStatus from Appointment order by AppointmentTime"
    cur.execute(select_all_query)  
    query_result=cur.fetchall()                                  
    cur.close()
    return("Appointment Records:\n"+ str(query_result).replace("),", ")\n") ).replace("[","").replace("]","")+"\n"


def sqlite_get_hour_reserved_on_target_date(conn, email, appt_time):    
    cur=conn.cursor()     
    dt_appt_time = datetime.strptime(appt_time, '%Y-%m-%d') 
    
    query_get_reserved_hour_on_target_date="/* get the reserved hour on the target date */" \
                     +" select strftime('%H', AppointmentTime) " \
                     +" from Appointment  " \
                     +" where LoginEmail<>'"+email + "' and strftime('%Y-%m-%d', AppointmentTime)='"+ dt_appt_time.strftime("%Y-%m-%d")+"'"        
    cur.execute(query_get_reserved_hour_on_target_date) 
    query_result=cur.fetchone() 
    # print(query_result)     
    cur.close()    
    if query_result==None:
        return 0
    else:
        return(int(query_result[0]))

# def sqlite_delete_old_appointment(conn, appt_date):    
#     cur=conn.cursor()     
#     dt_appt_date = datetime.strptime(appt_date, '%Y-%m-%d') 
#     appt_week_start_date = str(dt_appt_date-timedelta(days=dt_appt_date.weekday()))  # Monday is the start date of a week        
#     ### delete old records which are not in the curent appointment week
#     delete_query="delete from Appointment "+\
#                   "where AppointmentTime <'" + appt_week_start_date+"'"
#     print(delete_query)
#     cur.execute(delete_query);   
#     conn.commit()           
#     cur.close()


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
        return True
    else:
        return False


def sqlite_check_email_usability(conn, email, appt_time, court_number):    
    cur=conn.cursor()     
    dt_appt_time = datetime.strptime(appt_time, '%Y-%m-%d') 
    appt_week_start_date = str(dt_appt_time-timedelta(days=dt_appt_time.weekday()))  # Monday is the start date of a week
    # print(appt_week_start_date)    
    #### check whether the email has been used more than 3 times in the booking week, it doesn't matter on which date or court.   
    check_email_query="/* Check whether email has been used once in a day */" \
                    +" select LoginEmail " \
                    +" from Appointment  " \
                    +" where LoginEmail='"+email + "' and strftime('%Y-%m-%d', AppointmentTime)='"+ dt_appt_time.strftime("%Y-%m-%d")+"'" \
                    +" union "\
                    +"/* Check whether email status is ReachedDailyLimit1 */" \
                    +" select LoginEmail " \
                    +" from Appointment  " \
                    +" where AppointmentStatus='ReachedDailyLimit1' and LoginEmail='"+email +"' and strftime('%Y-%m-%d', AppointmentTime)='"+ dt_appt_time.strftime("%Y-%m-%d")+"'" \
                    +" union "\
                    +"/* Check whether email status is ReachedWeeklyLimit3 */" \
                    +" select LoginEmail " \
                    +" from Appointment  " \
                    +" where AppointmentStatus='ReachedWeeklyLimit3' and LoginEmail='"+email +"' and AppointmentTime>='" + appt_week_start_date +"'" \
                    +" union "  \
                    +"/* Check whether email has been used 3 times in a week*/" \
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
        return True
    else:
        return False   

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
# court_number="3"
# str_military_hour_option=""    
# # # str_military_hour_option = "[5,6]"    

##########################################################################################################
###### If enable_purge_record =True, Appointment table will be purged. This variable is set manually #####
###### Purge will be done by LoadTennisReservation.py in advance
##########################################################################################################
enable_purge_record=False


host_name=socket.gethostname()
###### login email list
# list_email=["xinbo.liu@gmail.com","liuxinbo.utube@gmail.com","xiewanqing2019@gmail.com"]


######## To limite the confliction 
if host_name=="MaoDesktop":
    list_email=["xinbo.liu@gmail.com","xiewanqing2019@gmail.com"]  
elif host_name=="mydesktop":
    list_email=["liuxinbo.utube@gmail.com"]
else: 
    list_email=["xinbo.liu@gmail.com","liuxinbo.utube@gmail.com","xiewanqing2019@gmail.com"]    
    
    
# ######## One password for all login emails   
# login_password="COUdl@1125"

######## Convert court_label to court_label 
if (court_number=="2"):
    court_label="Court #2 Arrive Residents Only"
elif (court_number=="3"):
    court_label="Court #3 Arrive Residents Only"
else:
    raise CourtNumberNotExists() 

 
######## Generate batch id and msg_summary
batch_id=int(datetime.now().strftime("%Y%m%d%H%M%S"))
msg_summary="\n\n"+"HostName: "+host_name + "\n" 
msg_summary=msg_summary+"BatchID: "+str(batch_id) + "\n" 


#################### Static Input Parameter: Program Configuration ##########################
##### Log file and db file are put in Data folder
# log_path=r"C:\Users\Me\Projects\GitHub\Utility\BookTennisCourt\Data\CourtReservationLog.txt"
db_file_path=r"C:\Users\Me\Projects\GitHub\Utility\BookTennisCourt\Data\CourtReservation.db"

##### Chrome driver is put in project code folder
chrome_driver_path = r"C:\Users\Me\Projects\GitHub\Utility\BookTennisCourt\chromedriver.exe" 

profile_directory=r"–user-data-dir=C:\Users\Me\AppData\Local\Google\Chrome\User Data"
login_url="https://app.courtreserve.com/Online/Account/LogIn/7629"

create_table_script="""CREATE TABLE IF NOT EXISTS Appointment
                       (
                          BatchID              BIGINT,
                          LoginEmail           VARCHAR(50) NOT NULL,
                          LoginTime            TEXT,
                          AppointmentTime      TEXT,
                          CourtNumber          INT,
                          Description          VARCHAR(50),
                          AppointmentStatus    VARCHAR(50)
                       )
                 """

try: 
    ######## create a sqlite db and table to track the usage of login_email  
 
    conn = create_tennis_court_booking_db_connection(db_file_path)
    create_appointment_table(conn, create_table_script) 
    
    ######## Purge Appointment table           
    if enable_purge_record==True:        
        sqlite_purge_appointment(conn)

    ######## Test record 
    ### sqlite_insert_appointment(conn, 111111, 'waynelxb@gmail.com', '2022-2-14', '2022-2-21', 3,'', 'ReachedWeeklyLimit3') 
    ######## At the end of the while loop, if is_email_usable=False, another login_email will be used.    
    
    is_email_usable=False
    while is_email_usable==False:   
        ##### Get login time 
        now=datetime.now()
        str_login_time=now.strftime("%Y-%m-%d %H:%M:%S") 
        # attribute_current_date_short_date=now.strftime("%#m/%#d/%Y") 
        # print(attribute_current_date_short_date)       

        #### The courts on the date which is 7 days later than the current date are released at 12pm. 
        str_court_release_time=now.strftime("%Y-%m-%d")+" "+"12:00:00"
        dt_login_time=datetime.strptime(str_login_time, "%Y-%m-%d %H:%M:%S")
        dt_court_release_time=datetime.strptime(str_court_release_time, "%Y-%m-%d %H:%M:%S")
        
        # if dt_login_time >= dt_court_release_time:
        #     dt_target_date=(now + timedelta(days=7)).date()
        # else:
        #     dt_target_date=(now + timedelta(days=6)).date() 
        
        ###### only target the date 7 days later
        dt_target_date=(now + timedelta(days=7)).date()

        str_target_date=str(dt_target_date) 
        # xpath_element_button_target_date="//a[@tabindex='-1'][@class='k-link'][@title='Friday, April 1, 2022']"                                                                                        
        attribute_formatted_target_date=dt_target_date.strftime("%A, %B %#d, %Y")   
        # print(attribute_formatted_target_date)
        str_target_day_of_month=dt_target_date.strftime("%#d")        
        # print(str_target_day_of_month)
   

        ####### If the length of str_military_hour_option is longer than 2, then conver it to a list, or use the default hour list   
        input_option_length=len(str_military_hour_option)        
        if(input_option_length>2):             
            list_military_hour_option=str_military_hour_option[1:input_option_length-1].split(',')
        else:
            #### use the default hour list    
            if(dt_target_date.weekday()>=5):
                list_military_hour_option=[15,16,17,18,13,14,10,11,12]
            else:
                list_military_hour_option=[17,18,16,19,20]                   
        #### Delete the old records not in the current appoitment week  
        # sqlite_delete_old_appointment(conn, str_target_date)

        #### Check whehter the court has been logged as overused  
        if sqlite_check_court_availability(conn, court_number, str_target_date)==False:
            raise CourtOverbooked()      
            
            
        msg_summary=msg_summary+"Login Time: "+str_login_time + "\n"   
        msg_summary=msg_summary+"Court Name: " + court_label +"\n"
        msg_summary=msg_summary+"Target Date: " + str_target_date +"\n"        
        msg_summary=msg_summary+"Expected Hour List: "+str(list_military_hour_option)+"\n"
        # print(msg_summary)            
            
        
        #### Find usable login email   
        login_email=""              
        for email in list_email:
            # print(email)
            if sqlite_check_email_usability(conn, email, str_target_date, court_number)  == True:
                login_email=email  
                if login_email=="xinbo.liu@gmail.com":
                    user_name="Xinbo" 
                    login_password="COUdl@1125"
                if login_email=="liuxinbo.utube@gmail.com":
                    user_name="Utube" 
                    login_password="COUdl@1125"
                if login_email=="xiewanqing2019@gmail.com":
                    user_name="Wanqing" 
                    login_password="1991@Qing"                   
                is_email_usable=True
                break   
        #### all emails are overused
        if is_email_usable==False:
            raise EmailNotUsable()

    
        msg_summary=msg_summary+"Login Email: "+login_email+"\n"
        # print(msg_summary)


        ####### Create chrome driver       
        svc=Service(chrome_driver_path)
        options = webdriver.ChromeOptions()
        options.add_argument(profile_directory)
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_experimental_option("excludeSwitches", ["enable-logging"])       
        driver = webdriver.Chrome(service=svc, options=options)       
        # driver.implicitly_wait(3) ## seconds 
        
        ######## Open Login page
        driver.get(login_url)        
        #### input email       
        id_element_email_inputbox="UserNameOrEmail" 
        element_email_inputbox=get_element_wait_for_load(1,"ID",id_element_email_inputbox)      
        element_email_inputbox.send_keys(login_email)      
        #### Input password 
        id_element_password_inputbox="Password"        
        element_password_inputbox=get_element_wait_for_load(1,"ID",id_element_password_inputbox)         
        element_password_inputbox.send_keys(login_password)     
        ### Click button               
        xpath_element_password_submit_button="//button[@class='btn btn-log btn-block btn-thm btn-submit']"
        element_password_submit_button=get_element_wait_for_load(1,"XPATH",xpath_element_password_submit_button) 
        element_password_submit_button.click()        
              
        ####### If current time is in (11:59:30, 12:00:00), then wait, or proceed
        while datetime.now()>dt_court_release_time + timedelta(seconds=-30) and datetime.now()<dt_court_release_time:
            time.sleep(1)
            
        
        msg_summary=msg_summary+"Booking Start Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
        
        ######## Switch to Announcements page        
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))        
        driver.switch_to.window(driver.window_handles[-1])
        #### Click Arrive Streeterville
        xpath_element_ArriveStreeterville="//li[@data-sm-show='true']/a[text()='Arrive Streeterville ']"      
        element_ArriveStreeterville=get_element_wait_for_load(1,"XPATH",xpath_element_ArriveStreeterville)          
        element_ArriveStreeterville.click()   

        
        ######## Switch to Schedule page       
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))        
        driver.switch_to.window(driver.window_handles[-1])    
        
        #<span class="k-icon k-i-calendar">
        xpath_element_button_calendar="//span[@class='k-icon k-i-calendar']"
        element_button_calendar=get_element_wait_for_load(1,"XPATH",xpath_element_button_calendar)
        element_button_calendar.click()
        time.sleep(0.5)
        
        
        ### Click target date
        # <a tabindex="-1" class="k-link" href="#" data-value="2022/3/1" title="Friday, April 1, 2022">1</a> 
        # xpath_element_button_target_date="//a[@tabindex='-1'][@class='k-link'][@title='"+ attribute_formatted_target_date +"']"     
        
        xpath_element_button_target_date="//a[@tabindex='-1'][@class='k-link'][text()='"+ str_target_day_of_month +"']"
        
        #### if cannot find the date, then sleep 2 second then refresh, if still cannot find it, then raise exception
        if get_element_wait_for_load(1,"XPATH",xpath_element_button_target_date)=="None":
            time.sleep(1)
            driver.refresh()
            if get_element_wait_for_load(1,"XPATH",xpath_element_button_target_date)=="None":
                raise ElementLocatorNotExists("xpath_element_button_target_date="+ xpath_element_button_target_date)               
        element_button_target_date=get_element_wait_for_load(1,"XPATH",xpath_element_button_target_date)
        element_button_target_date.click()
        
        
        ######## Switch to the target date page   
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))        
        driver.switch_to.window(driver.window_handles[-1])  
        
        msg_summary=msg_summary+"Check Hour Start Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
        
        ####### Get the hour has been reserved on the target date 
        reserved_hour_on_target_date=sqlite_get_hour_reserved_on_target_date(conn, login_email, str_target_date)       
        ### Find and click available target hour
        is_target_hour_available=False 
        for target_military_hour in list_military_hour_option: 
            ###### Try the hour which is not the same hour already reserved on the target date
            if target_military_hour!=reserved_hour_on_target_date:
                target_hour=("0"+str(target_military_hour))[-2:]+":00:00"
                attribute_formatted_target_date_hour=dt_target_date.strftime("%a %b %d %Y") +" "+ target_hour 
                ### str_target_date_hour is used in database
                str_target_date_hour=dt_target_date.strftime("%Y-%m-%d") +" "+ target_hour   
                # print(str_target_date_hour)                       
                # <button start="Fri Apr 01 2022 08:00:00 GMT-0500 (Central Daylight Time)" end="Fri Apr 01 2022 09:00:00 GMT-0500 (Central Daylight Time)" instructorid="undefined" courtlabel="Court #3 Arrive Residents Only" class="btn btn-default slot-btn m-auto">Reserve 8:00 AM</button>
                xpath_element_button_target_date_court_time="//button[contains(@start, '"+attribute_formatted_target_date_hour +"')][@courtlabel='"+court_label+"'][@class='btn btn-default slot-btn m-auto']"         
                # print(xpath_element_button_target_date_court_time)                       
                if get_element_wait_for_load(0.5,"XPATH", xpath_element_button_target_date_court_time) != "None":                
                    element_button_target_date_court_time=get_element_wait_for_load(0.3,"XPATH",xpath_element_button_target_date_court_time)                 
                    element_button_target_date_court_time.click() 
                    time.sleep(0.2)
                    msg_summary=msg_summary+attribute_formatted_target_date_hour +" is available for court "+ str(court_number)+"\n"
                    is_target_hour_available=True
                    break
                else: 
                    msg_summary=msg_summary+attribute_formatted_target_date_hour+" is NOT available for court "+ str(court_number)+"\n"           
                
        if is_target_hour_available==False:
            raise TimeNotAvailable()                   
         
        msg_summary=msg_summary+"Check Hour End Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
        
        ######## Swith to Player page
        driver.switch_to.window(driver.window_handles[0])
        
        # ###### If Important Message Page appears, the emai has been overused in a day or week
        # # <div class="modal-body">You have reached max number of courts allowed to reserve per day: 1</div>             
        # xpath_element_important_message_page="//div[contains(text(),'You have reached max number of courts allowed to reserve per day: 1')]"       
        # if get_element_wait_for_load(0.5,"XPATH",xpath_element_important_message_page)!="None":
        #     sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_target_date, court_number, '', "ReachedDailyLimit1")
        #     driver.quit()
        #     raise EmailNotUsable()
        # #<div class="modal-body">You have reached max number of courts allowed to reserve per week: 3</div>
        # xpath_element_important_message_page="//div[contains(text(),'You have reached max number of courts allowed to reserve per week: 3')]"  
        # if get_element_wait_for_load(0.5,"XPATH",xpath_element_important_message_page)!="None":
        #     sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_target_date, court_number, '', "ReachedWeeklyLimit3")
        #     driver.quit()
        #     raise EmailNotUsable()

        
        # <textarea autocomplete="off" class="required form-control" id="_0__Value" name="Udfs[0].Value"></textarea>          
        xpath_element_textarea_resident_with_you="//textarea[@autocomplete='off'][@class='required form-control']"      
        element_textarea_resident_with_you=get_element_wait_for_load(5,"XPATH",xpath_element_textarea_resident_with_you)
        if element_textarea_resident_with_you=="None":
            raise ElementLocatorNotExists("xpath_element_textarea_resident_with_you="+ xpath_element_textarea_resident_with_you)          
        element_textarea_resident_with_you.send_keys("Jiajia Guo") 
        time.sleep(0.2)
        

        ###### Click the footer SAVE button        
        # <div class="modal-footer-container"><div class="modal-title-buttons"><button type="reset" class="btn btn-light" data-dismiss="modal">Close</button><button type="button" class="btn btn-primary btn-submit " onclick="submitCreateReservationForm()">Save</button></div></div>
        xpath_element_bottom_save_button="//div[@class='modal-footer-container']//button[@type='button'][text()='Save']"
        element_bottom_save_button=get_element_wait_for_load(1,"XPATH",xpath_element_bottom_save_button)
        if element_bottom_save_button=="None":
            raise ElementLocatorNotExists("xpath_element_bottom_save_button="+ xpath_element_bottom_save_button)          
        element_bottom_save_button.click() 
        time.sleep(0.2)
        
        
        ###### Switch to Close page        
        driver.switch_to.window(driver.window_handles[0])        
        # <button type="reset" class="btn btn-light" data-dismiss="modal">Close</button>
        xpath_element_close_button="//button[@type='reset'][@data-dismiss='modal'][text()='Close']"
        element_close_button=get_element_wait_for_load(1,"XPATH",xpath_element_close_button)
        if element_close_button=="None":
            raise ElementLocatorNotExists("xpath_element_close_button="+ xpath_element_close_button)  
        element_close_button.click()        
        msg_summary=msg_summary+"Booking End Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
        # time.sleep(3)        

        # # ###### Switch to Reservation page                 
        # WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))        
        # driver.switch_to.window(driver.window_handles[-1]) 
        # time.sleep(3)
        # driver.refresh()
        # # xpath_element_reserved_datetime_container="//div[@style='display:block;padding-top:2px;']/a[@class='btn-scheduler-edit-details']"   
        # xpath_element_reserved_datetime_container="//a[@class='btn-scheduler-edit-details'][text()='Details']"   
        # element_reserved_datetime_container=get_element_wait_for_load(3,"XPATH",xpath_element_reserved_datetime_container) 
        # if element_reserved_datetime_container=="None":
        #     raise ElementLocatorNotExists("xpath_element_reserved_datetime_container="+ xpath_element_reserved_datetime_container)   
        # element_reserved_datetime_container.click()        
        # time.sleep(3)      
    
        # ###### Switch to Reservation Details page        
        # driver.switch_to.window(driver.window_handles[0])        
        # #<span class="title-part">5481#</span>
        # xpath_element_door_code="//span[@class='title-part'][contains(text(),'#')][not(contains(text(),'Court'))]" 
        # # door_code=get_element_wait_for_load(3,"XPATH",xpath_element_door_code).text             
        # door_code=driver.find_element(By.XPATH, xpath_element_door_code).text  
        # sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_target_date_hour, court_number,"WeekDay: "+dt_target_date.strftime('%a')+" | Code: "+door_code, "Succeeded")  
        
        sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str_target_date_hour, court_number,"WeekDay: "+dt_target_date.strftime('%a'), "Succeeded")      
        driver.quit()                                
        
        #### Create calendar event
        # calendar_event_status=create_calendar_event(dt_target_date, target_military_hour, "Court "+str(court_number)+" Code: "+ door_code +" Login:"+ login_email)
        calendar_event_status=create_calendar_event(dt_target_date, target_military_hour, "Court "+str(court_number)+" Login:"+ login_email)
        
        #### Send success email
        msg_summary=msg_summary+"Reservations:"+sqlite_get_appointment(conn)+calendar_event_status+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"        
        send_email("Tennis Court Booking Succeeded", msg_summary) 
                   

except CourtOverbooked:
    msg_summary=msg_summary+"Exception: " +court_label+" has been overbooked.\n"+ sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    print(msg_summary)
    # log_process(log_path, msg_summary)    
    send_email("Tennis Court Booking Failed", msg_summary)   

except EmailNotUsable:
    msg_summary=msg_summary+"Exception: One or all the emails in "+str(list_email)+" have been overused for target date "+str_target_date+"!\n"+sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    print(msg_summary)
    # log_process(log_path, msg_summary)    
    send_email("Tennis Court Booking Failed", msg_summary)    
    
except ElementLocatorNotExists as e:
    driver.quit()
    msg_summary=msg_summary+"Error: "+ e.message +" cannot be loacated.\n"+sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    print(msg_summary)
    # log_process(log_path, msg_summary)    
    send_email("Tennis Court Booking Failed", msg_summary)   
    
except CourtNumberNotExists:
    msg_summary=msg_summary+"Error: Court "+ court_number + " doesn't exist.\n"+sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    print(msg_summary)
    # log_process(log_path, msg_summary)
    send_email("Tennis Court Booking Failed", msg_summary)   
    
except TimeNotAvailable:
    driver.quit()
    msg_summary=msg_summary+"Error: No time slot in "+ str(list_military_hour_option) + " is available.\n"+sqlite_get_appointment(conn) +"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    print(msg_summary)
    # log_process(log_path, msg_summary)
    send_email("Tennis Court Booking Failed", msg_summary)    

except: 
    driver.quit()    
    exc_type, exc_value, exc_traceback = sys.exc_info()
    exceptMessage=repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
    msg_summary=msg_summary+"Error: "+exceptMessage +"\n" +sqlite_get_appointment(conn)+"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    print(msg_summary)   
    # log_process(log_path, msg_summary)
    send_email("Tennis Court Booking Failed", msg_summary)    