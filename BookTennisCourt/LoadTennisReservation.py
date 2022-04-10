###########################
##### Version 2.0 #########
###########################

# import os
import time
import sys
import traceback
from datetime import datetime  
# from datetime import timedelta  
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
# from cal_setup import get_calendar_service
import re



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
    select_all_query="select * from Appointment"
    cur.execute(select_all_query)  
    query_result=cur.fetchall()                                  
    cur.close()
    return("Appointment Records:\n"+ str(query_result).replace("),", ")\n") ).replace("[","").replace("]","")+"\n"

######## Get Info from Reservation Table ##########
class WebTable:
    def __init__(self, webtable):
       self.table = webtable

    def get_row_count(self):
      return len(self.table.find_elements(By.TAG_NAME, "tr") )- 1

    def get_column_count(self):
        return len(self.table.find_elements(By.XPATH,"//tr[2]/td"))

    def get_table_size(self):
        return {"rows": self.get_row_count(),
                "columns": self.get_column_count()}
    
    def row_data(self, row_number):
        # if(row_number == 0):
        #     raise Exception("Row number starts from 1")
        
        row_number = row_number + 1
        row = self.table.find_elements(By.XPATH,"//tr["+str(row_number)+"]/td")
        rData = []
        for webElement in row :
            rData.append(webElement.text)
        return rData    
    
    def column_data(self, column_number):
        col = self.table.find_elements(By.XPATH,"//tr/td["+str(column_number)+"]")
        rData = []
        for webElement in col :
            rData.append(webElement.text)
        return rData

    def get_all_data(self):
        # get number of rows
        noOfRows = len(self.table.find_elements(By.XPATH,"//tr")) -1
        # get number of columns
        noOfColumns = len(self.table.find_elements(By.XPATH,"//tr[1]/td"))
        allData = []
        # iterate over the rows, to ignore the headers we have started the i with '1'
        for i in range(2, noOfRows):
            # reset the row data every time
            ro = []
            # iterate over columns
            for j in range(1, noOfColumns) :
                # get text from the i th row and j th column
                ro.append(self.table.find_element(By.XPATH,"//tr["+str(i)+"]/td["+str(j)+"]").text)

            # add the row data to allData of the self.table
            allData.append(ro)
        return allData

    def presence_of_data(self, data):

        # verify the data by getting the size of the element matches based on the text/data passed
        dataSize = len(self.table.find_elements(By.XPATH,"//td[normalize-space(text())='"+data+"']"))
        presence = False
        if(dataSize > 0):
            presence = True
        return presence

    def get_cell_data(self, row_number, column_number):
        if(row_number == 0):
            raise Exception("Row number starts from 1")
        row_number = row_number+1
        cellData = self.table.find_element(By.XPATH,"//tr["+str(row_number)+"]/td["+str(column_number)+"]").text
        return cellData


##########################################################################################################
###### If enable_purge_record =True, Appointment table will be purged. This variable is set manually #####
###### Always purge the table to get the latest existing reservations
##########################################################################################################
enable_purge_record=True


###### login email list"xinbo.liu@gmail.com"
list_email=["xinbo.liu@gmail.com","liuxinbo.utube@gmail.com"]
######## One password for all login emails     
login_password="COUdl@1125"
 
######## Generate batch id and msg_summary
batch_id=int(datetime.now().strftime("%Y%m%d%H%M%S"))
msg_summary="\n\n"+"BatchID: "+str(batch_id) + "\n"   

      
#################### Static Input Parameter: Program Configuration ##########################
##### Log file and db file are put in Data folder
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
  
    for login_email in list_email:       
        now=datetime.now()
        str_login_time=now.strftime("%Y-%m-%d %H:%M:%S") 
        attribute_current_date_short_date=now.strftime("%#m/%#d/%Y")      
        taget_year = now.year            
        
        ####### Create chrome driver       
        svc=Service(chrome_driver_path)
        options = webdriver.ChromeOptions()
        options.add_argument(profile_directory)
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_experimental_option("excludeSwitches", ["enable-logging"])       
        driver = webdriver.Chrome(service=svc, options=options)        
        
        
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
        
        
        ####### Open Reservation page to load existing reservations into the sqlite table
        driver.get("https://app.courtreserve.com/Online/MyProfile/MyReservations/7629?page=upcoming")    
        time.sleep(5)          
        xpath_element_reservation_table="//table[@role='grid']"
        w = WebTable(get_element_wait_for_load(5,"XPATH",xpath_element_reservation_table))            
        # print("First row data : ", w.row_data(0))          
        for i in range(5):
            if len(w.row_data(i))>0:    
                str_reserved_time=w.row_data(i)[3]
                dash_index=str_reserved_time.find("-")
                str_reserved_time=str_reserved_time[:dash_index]             
                str_reserved_datetime = w.row_data(i)[2]+" "+str(taget_year) + " "+ str_reserved_time    
                str_reserved_datetime=str_reserved_datetime.replace("st","").replace("nd","").replace("rd","").replace("th","")
                comma_index=str_reserved_datetime.find(",")
                str_reserved_datetime=str_reserved_datetime[comma_index+1:].strip()
                dt_reserved_datetime=datetime.strptime(str_reserved_datetime, "%b %d %Y %I:%M %p")                   
                # print(dt_reserved_datetime)               
                str_reserved_court=w.row_data(i)[5]
                reservered_court_number=re.findall(r'\d+', str_reserved_court)       
                # print(reservered_court_number[0])                   
                str_reserved_code=w.row_data(i)[6]
                # print(str_reserved_code)        
                sqlite_insert_appointment(conn, batch_id, login_email, str_login_time, str(dt_reserved_datetime), reservered_court_number[0], "WeekDay: "+dt_reserved_datetime.strftime('%a')+" | Code: "+str_reserved_code,"Succeeded")
        driver.quit()          
        
    print(sqlite_get_appointment(conn)) 
    #send_email("Tennis Court Reservation Load Succeeded", sqlite_get_appointment(conn))    
except: 
    driver.quit()    
    exc_type, exc_value, exc_traceback = sys.exc_info()
    exceptMessage=repr(traceback.format_exception(exc_type, exc_value, exc_traceback))
    msg_summary=msg_summary+"Error: "+exceptMessage +"\n" +"Logout Time: "+datetime.now().strftime("%Y-%m-%d %H:%M:%S")+"\n"
    # print(msg_summary)   
    send_email("Tennis Court Reservation Load Failed", msg_summary)