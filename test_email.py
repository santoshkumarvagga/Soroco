"""
This module does some calcualtion over excel sheet sent through mail and reverts back with the calculated result.

pre-requisites:
* The excel sheet attached mail should already be sent/present to/in sender/source mail id
* Speed of network connection should not be too slow, in order not to exceed given wait time periods in code.
* manually authenticate for first time gmail login(using phone notification) on a new connection

NOTE: The code retries the process for 1 more time, if desired end point is not reached in first attempt.

"""
import xlrd
import unittest
import smtplib
import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# define logging 
logging.basicConfig(filename='log_email_gmail.txt',format='%(asctime)s : %(filename)s : %(funcName)s : %(levelname)s :  %(lineno)d - %(message)s', level = logging.DEBUG)

class TestEmailAutomation(unittest.TestCase):
    '''Contains two test cases to illustrate excel sheet automation via gmail'''

    SENDER_MAIL = 'vaggasantoshkumar@gmail.com'
    SENDER_PASSWORD = 'Santosh@96'

    RECEIVER_MAIL = 'vaggasantoshkumar@gmail.com'

    flag = False
    fetch_success = False

    # opening log file for fresh update
    logfile = open('log_email_gmail.txt','w')
    # identidfier for fresh excetion start point in log file
    logging.critical('<---------------Fresh Exection Started Now----------------->')


    def test_get_sheet(self):
        '''
           This method logs into the gmail account of sender fetches the attached excel sheet from the corresponding inbox mail
           and gives it for needed calcualtion.
           Later, the result is sent to recipient mail.
        '''
        try: 
            driver = webdriver.Chrome()

            logging.info("Gmail login started..")

            driver.get(r'https://accounts.google.com/signin/v2/identifier?continue='+\
            'https%3A%2F%2Fmail.google.com%2Fmail%2F&service=mail&sacu=1&rip=1'+\
            '&flowName=GlifWebSignIn&flowEntry = ServiceLogin') 
            driver.implicitly_wait(15) 
    
            loginBox = driver.find_element_by_xpath('//*[@id ="identifierId"]') 
            loginBox.send_keys(self.SENDER_MAIL) 
    
            nextButton = driver.find_elements_by_xpath('//*[@id ="identifierNext"]') 
            nextButton[0].click() 
    
            passWordBox = driver.find_element_by_xpath( 
    	        '//*[@id ="password"]/div[1]/div / div[1]/input')
            passWordBox.send_keys(self.SENDER_PASSWORD) 
    
            nextButton = driver.find_elements_by_xpath('//*[@id ="passwordNext"]')
            nextButton[0].click() 
    
            logging.debug('Gmail Login Successful...!!') 
        except: 
            logging.warning('Gmail Login Failed') 

        logging.debug("waiting for 100 sec..")
        driver.implicitly_wait(100)
        search = driver.find_element_by_xpath('//*[@id="gs_lc50"]/input[1]')

        # seacrh user mail account using the SUBJECT of the sender mail as query string
        search.send_keys('Time sheet details review')
        search.send_keys(Keys.ENTER)

        logging.debug('Located latest sender email with excel sheet.')

        # wait untill page loads completely
        logging.debug("waiting for 60 sec..")
        driver.implicitly_wait(60)

        inbox_mail = driver.find_element_by_xpath('//*[@id=":nb"]/span')
        inbox_mail.click()

        # wait untill page loads completely
        logging.debug("waiting for 60 sec..")
        driver.implicitly_wait(60)

        # NOTE: It's possible to fetch attached excel sheet url address via selenium, but i have failed to do so with multiple failure attempts'
        # hence embedding the url of the excel sheet directly to download it to Downloads folder.

        # download url for excel sheet in sender mail
        try:
            driver.get('https://mail.google.com/mail/u/0/https://mail.google.com/mail/u/0?ui=2&ik=203a6d8d44&attid=0.1&permmsgid=msg-a:r-6768357511561605125&th=1742e4bbe0e07bac&view=att&disp=safe&realattid=f_keccbu160')
        except exception as e:
            logging.error('{} raised. Failed to download excel sheet from sender mail.'.format(e))
        else:
            logging.debug('Downlaod Started..!!')
            logging.info('Successfully downloaded latest excel sheet!!')
            self.fetch_success = True
    
    def test_send_reply(self):
        '''
            This method takes input excel sheet
            and finds the total work hours from it and sends mail to recepient
         '''
        if True:
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

            s.starttls()

            # login to sender gmail account
            s.login(self.SENDER_MAIL, self.SENDER_PASSWORD)
            
            # display the expected result
            message = "Calculated the total work hours - {} hours!!!".format(sum)
            
            # send mail to receiver mail id
            try:
                s.sendmail(self.SENDER_MAIL, [self.RECEIVER_MAIL], message)
            except Exception as e:
                logging.error('{} raised, Failed to send the response mail to recipient.'.format(e))
                assert False, 'Failed to send the response mail to recipient.'
            else:
                logging.info('Successfully sent result to recipient mail. Please Check recipient mail..!!')
            s.quit()
            self.logfile.close()

if __name__ == '__main__':
    unittest.main()
