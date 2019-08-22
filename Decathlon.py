import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import logging
import traceback
import os
import pdb
import time
import sys
from progress.bar import Bar


#N.Sathishkumar@expeditors.com
#123456


class downloader:
    
    def __init__(self):
        
        self.logger = logging.getLogger('Cape System Log')
        self.logger.setLevel(logging.DEBUG)
        
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        
        self.logger.addHandler(ch)
        self.logger.propagate = False
        #self.bar = Bar('Processing', max=10)

        self.consolidationDiv = "//*[@id='menu_li_4']"
        self.consolButton = "//*[@id='actions']/li[3]/img"
        self.filterButton = "//*[@id='filters-action-full']"
        self.poEntryInput = "//*[@id='flatGroup']/li[3]/div/div[3]/input"

        # input settings
        self.inputFile = r"D:\CLR.xlsx"
        self.inputs = pd.read_excel(self.inputFile)
        #print(self.inputs)
        
        #Variables
        self.orderNumbers = ''

        # application setting
        
        self.url = "https://cape-asia.decathlon.net/capetm/index.jsp"
        
    def setupChrome(self):

        self.logger.info("Setup Chrome function called")

        # Contains all chrome settings
        self.logger.info("Setting-up Chrome")
        self.settings = webdriver.ChromeOptions()
        self.settings.add_argument("--incognito")
        self.settings.add_argument('--ignore-ssl-errors')
        self.settings.add_argument('--ignore-certificate-errors')
        self.settings.add_argument('–-disable-web-security')
        self.settings.add_argument('–-allow-running-insecure-content')
        #self.settings.add_argument('--browser.download.folderList=2')
        #self.settings.add_argument('--browser.helperApps.neverAsk.saveToDisk=text/csv')
        #self.settings.add_experimental_option("prefs",profile)

        #self.bar.next()
            
        
    def loadBrowser(self):
        
        self.logger.info("Load Browser function called")
        #pdb.set_trace()
        self.setupChrome()

        try:
            #self.browser = webdriver.Chrome("D:\\DataScrapping\\ProjectBigSchedules\\chromedriver.exe")
            self.browser = webdriver.Chrome(chrome_options=self.settings, executable_path=r"D:\chromedriver.exe")
            self.browser.maximize_window()
            self.logger.info("Page setup complete will now go to the URL")
            #self.bar = next()

        except Exception as e:
            self.logger.critical("Unable to load chrome driver. " + str(e))
        
        #Entering the URLc

        
        self.browser.get(self.url)

        delay = 60 # seconds
        try:
            myElem = WebDriverWait(self.browser, delay).until(EC.presence_of_element_located((By.ID, 'username')))
            self.logger.info("Page is ready!")
            
            
        except TimeoutException:
            self.logger.info("Loading took too much time! Exiting now")
            sys.exit(0)

        inputElement = self.browser.find_element_by_id("username")
        inputElement.send_keys('N.Sathishkumar@expeditors.com')

        passwardField = self.browser.find_element_by_id('password')
        passwardField.send_keys('123456')

        submitButton = self.browser.find_element_by_xpath("/html/body/div[1]/div/form/div[4]/button")
        submitButton.click()

        #3
        #self.bar.next()

        #pdb.set_trace()
        self.asOfConsolidation()

    def wait_for_class_to_be_available(self,browser,elementXpath, total_wait=100):
        try:
            element = WebDriverWait(self.browser, 15).until(
                EC.element_to_be_clickable((By.XPATH, elementXpath)))
            element.click()
        except Exception as e:
            print("Wait Timed out")
            print(e)
            total_wait -= 1
            time.sleep(1)
            if total_wait > 1: 
                self.wait_for_class_to_be_available(self.browser,elementXpath, total_wait)

    def asOfConsolidation(self):
        
        #pdb.set_trace()
        self.logger.info("Consolidation function called")

        #clicking the consolidation div
        self.wait_for_class_to_be_available(self.browser,self.consolidationDiv)

        #clicking the consolidation button

        time.sleep(2)
        self.wait_for_class_to_be_available(self.browser,self.consolButton)

        #clicking the filterOption 
        time.sleep(2)
        self.wait_for_class_to_be_available(self.browser,self.filterButton)

        enterInput = self.browser.find_element_by_xpath(self.poEntryInput)
        #enterInput.send_keys(' '.join(self.poNumbers))

        enterInput.send_keys(self.orderNumbers)

        #Click the tickButton
        tickButton = self.browser.find_element_by_xpath("//*[@id='flatGroup']/li[3]/div/div[4]/img")
        #tickButton.click()

        #4
        #self.bar.next()

    def iterateOverInput(self):
        
        #print(self.inputs)

        self.orderNumbers = self.inputs['ORDER NUMBER'].values.tolist()
        print(self.orderNumbers)
        
        self.orderNumbers = ' '.join(str(e) for e in self.orderNumbers)
        

if __name__ == '__main__':
    obj = downloader()
    obj.iterateOverInput()
    obj.loadBrowser()