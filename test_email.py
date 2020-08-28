import smtplib
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlrd
import urllib3

options = webdriver.ChromeOptions() 
options.add_argument("download.default_directory='~\\Downloads\\'")

def send_reply():
    '''
        This method takes input excel sheet 
        and finds the total work hours from it ans sends mail to recepient
     '''
    
     # open the excel sheet
    workbook = xlrd.open_workbook('~\\Downloads\\Tracker.xlsx')

    # select the sheet
    sheet1 = workbook.sheet_by_name('Sheet1')

    # iterate throguh Hours coulmn
    entries = []
    for i in range(2,sheet1.nrows):
        entries.append(sheet1.cell(i,4).value)
    
    # find the total sum
    sum = 0
    for i in entries:
        if type(i)==float:
            sum += i

    # establish connection to gmail
    s = smtplib.SMTP('smtp.gmail.com', 587)

    # provide sender mail id and password
    SENDER_MAIL = "vaggasantoshkumar@gmail.com"
    SENDER_PASSWORD = "Santosh@96"

    # receiver mail id
    RECEIVER_MAIL = "vaggasantoshkumar@gmail.com"

    s.starttls()

    # login to sender gmail account
    s.login(SENDER_MAIL, SENDER_PASSWORD)

    # display the expected result
    message = "Calculated the total work hours - {} hours!!!".format(sum)

    # send mail to receiver mail id
    s.sendmail(SENDER_MAIL, [RECEIVER_MAIL], message)

    s.quit()


def get_sheet():
    '''
       This method takes the attached excel sheet from the sender mail 
       and gives it for needed calcualtion.
    '''
    gmailId = 'vaggasantoshkumar@gmail.com'
    passWord = 'Santosh@96'
    try: 
        driver = webdriver.Chrome(chrome_options=options)
        driver.get(r'https://accounts.google.com/signin/v2/identifier?continue='+\
        'https%3A%2F%2Fmail.google.com%2Fmail%2F&service=mail&sacu=1&rip=1'+\
        '&flowName=GlifWebSignIn&flowEntry = ServiceLogin') 
        driver.implicitly_wait(15) 
    
        loginBox = driver.find_element_by_xpath('//*[@id ="identifierId"]') 
        loginBox.send_keys(gmailId) 
    
        nextButton = driver.find_elements_by_xpath('//*[@id ="identifierNext"]') 
        nextButton[0].click() 
    
        passWordBox = driver.find_element_by_xpath( 
    	    '//*[@id ="password"]/div[1]/div / div[1]/input')
        passWordBox.send_keys(passWord) 
    
        nextButton = driver.find_elements_by_xpath('//*[@id ="passwordNext"]') 
        nextButton[0].click() 
    
        print('Login Successful...!!') 
    except: 
        print('Login Failed') 

    print("waiting for 100 sec..")
    driver.implicitly_wait(100)
    search = driver.find_element_by_xpath('//*[@id="gs_lc50"]/input[1]')
    search.send_keys('Time sheet details review')
    #search.click()
    search.send_keys(Keys.ENTER)
    
    print("waiting for 60 sec..")
    driver.implicitly_wait(60)

    inbox_mail = driver.find_element_by_xpath('//*[@id=":nb"]/span')
    inbox_mail.click()

    print("waiting for 60 sec..")
    driver.implicitly_wait(60)

    download_attachment = driver.find_element_by_xpath('//*[@id=":so"]')
    download_attachment.click()

    print("waiting for 10 sec..")
    driver.implicitly_wait(10)

    #locate_excel_file = driver.find_element_by_xpath('')
    #print(driver.current_url)

    excel_url = driver.find_element_by_xpath('//*[@id=":re"]').get_attribute('href')
    
    print("The download url is", excel_url)

    http = urllib3.PoolManager()

    http.request('GET', excel_url)

    print('Downlaod Started..!!')
    driver.implicitly_wait(50)

    print('Successfully downloaded latest excel sheet!!')

get_sheet()
send_reply()
