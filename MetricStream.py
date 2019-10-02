import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import time
import logging
import pdb
import shutil
import os


class PoTracker:
    
    def __init__(self):
        
        self.logger = logging.getLogger('GT-Nexus')
        self.logger.setLevel(logging.DEBUG)
        
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        
        self.logger.addHandler(ch)
        self.logger.propagate = False

        #xpaths

        self.loginLink = "//*[@id='loginHref']"
        self.approveLink = "//*[contains(text(), 'Approve Response')]"
        self.pagelink = "//*[@id='CmbID']"

        self.selectAll = "//div[contains(@class, 'x-combo-list-item') and text()='All']"
 

        
        self.url = "https://expeditors.metricstream.com/"
        
        
        # input settings
        
        
    def setupChrome(self):

        # Contains all chrome settings
        self.logger.info("Setting-up Chrome")
        self.settings = webdriver.ChromeOptions()
        #self.settings.add_argument("--incognito")
        self.settings.add_argument("--incognito")
        self.settings.add_argument('--ignore-ssl-errors')
        self.settings.add_argument('--ignore-certificate-errors')
        self.settings.add_argument('–-disable-web-security')
        self.settings.add_argument('–-allow-running-insecure-content')

            

    def wait_for_class_to_be_available(self,browser,elementXpath, total_wait=100):
        #pdb.set_trace()
        try:
            element = WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, elementXpath)))
            element.click()
        except Exception as e:
            print("Waiting for element to be Clickable")
            #print(e)
            total_wait = total_wait - 1
            time.sleep(1)
            if total_wait > 1: 
                self.wait_for_class_to_be_available(browser,elementXpath, total_wait)


    def loadBrowser(self):
        
        #pdb.set_trace()
        self.setupChrome()

        try:
            #self.browser = webdriver.Chrome("D:\\DataScrapping\\ProjectBigSchedules\\chromedriver.exe")
            self.browser = webdriver.Chrome(chrome_options=self.settings, executable_path=r"d:\Desktop\MAtrixStream\chromedriver.exe")
            self.browser.maximize_window()

        except Exception as e:
            self.logger.critical("Unable to load chrome driver. " + str(e))
        
        #Entering the URL
        self.logger.info("Getting URL")
        self.browser.get(self.url)


        self.logger.info("Looking for LoginLink")
        element = WebDriverWait(self.browser, 15).until(EC.element_to_be_clickable((By.XPATH, self.loginLink)))
        element.click()

        self.logger.info("Looking for Audit Link")

        audit = WebDriverWait(self.browser, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Audits and Inspections')]")))
        audit.click()

        approve = WebDriverWait(self.browser, 15).until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Approve Response')]")))
        approve.click()

        time.sleep(5)
        elems = self.browser.find_elements_by_xpath("//*[@id='ext-gen27-gp-6-Survey-bd']")
        for i in elems:

            all_li = i.find_elements_by_tag_name("a")
            for elem in all_li:
                print (elem.text)
            
        selectid = self.browser.find_element_by_xpath("//input[@id='CmbID']")
        selectid.click()

        
                                    
    def setupPage(self):

        self.setupChrome()
        self.loadBrowser()

if __name__ == '__main__':
    obj = PoTracker()
    obj.setupPage()