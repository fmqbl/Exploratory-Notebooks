import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
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
        self.tickButton = "//*[@id='flatGroup']/li[3]/div/div[4]/img"
        self.searchButton = "//*[@id='launchSearch']/img"
        self.tickAtBottomMenu = "//*[@id='callForConsoButton']/img"
        self.containerButton = "//*[@id='rightPartButtonAction']/a[1]/img"
        self.createContainer = "//*[@id='callContainerForm']/div[11]/button/span/em"
        
        #from input
        self.totalPieces = ""
        self.totalCtn = ""
        self.netWeight = ""
        self.grossWeight = ""
        self.totalVolume = ""

        #before selection datarow
        self.totalPiecesXPath = "//*[@id='globalDataLeft']/div/table/tbody/tr/td[3]/ul/li/label"
        self.totalCtnXPath = "//*[@id='globalDataLeft']/div/table/tbody/tr/td[4]/ul/li/label"
        self.netWeightXPath = "//*[@id='globalDataLeft']/div/table/tbody/tr/td[6]/ul/li/label"
        self.grossWeightXPath = "//*[@id='globalDataLeft']/div/table/tbody/tr/td[7]/ul/li/label"
        self.totalVolumeXPath = "//*[@id='globalDataLeft']/div/table/tbody/tr/td[8]/ul/li/label"
       
        #checkbox
        self.selectAllCheck = "//*[@id='mainAsCheckBox']"
        
        #Afterselection data row
        self.selectedPieces = "//*[@id='selectedDataLeft']/div/table/tbody/tr/td[3]/ul/li/label"
        self.selectedCtn = "//*[@id='selectedDataLeft']/div/table/tbody/tr/td[4]/ul/li/label"
        self.selectedNetWeight = "//*[@id='selectedDataLeft']/div/table/tbody/tr/td[6]/ul/li/label"
        self.selectedGrossWeight = "//*[@id='selectedDataLeft']/div/table/tbody/tr/td[7]/ul/li/label"
        self.selectedTotalVolume = "//*[@id='selectedDataLeft']/div/table/tbody/tr/td[8]/ul/li/label"

        #self.selectAllCheck = 

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
        
        #Entering the URL for validation of things 
        
        self.browser.get(self.url)

        delay = 40 # seconds
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
    #added comments
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

    def wait_for_xpath_to_present(self,browser,elementXpath, total_wait=100):
        #pdb.set_trace()
        try:
            element = WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.XPATH, elementXpath)))
            return element.text
        except Exception as e:
            print("Waiting for element to be Clickable")
            #print(e)
            total_wait = total_wait - 1
            time.sleep(1)
            if total_wait > 1: 
                self.wait_for_class_to_be_available(browser,elementXpath, total_wait)


    
    def asOfConsolidation(self):
        
        #pdb.set_trace()
        self.logger.info("Consolidation function called")

        #clicking the consolidation div
        #pdb.set_trace()
        self.browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        self.wait_for_class_to_be_available(self.browser,self.consolidationDiv)

        #clicking the consolidation button
        self.browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL + Keys.HOME)
        #time.sleep(2)
        self.wait_for_class_to_be_available(self.browser,self.consolButton)

        #clicking the filterOption 
        #time.sleep(2)
        self.wait_for_class_to_be_available(self.browser,self.filterButton)

        enterInput = self.browser.find_element_by_xpath(self.poEntryInput)
        #enterInput.send_keys(' '.join(self.poNumbers))
        
        enterInput.send_keys(self.orderNumbers)

        #Click the tickButton

        self.wait_for_class_to_be_available(self.browser,self.tickButton)

        #self.browser.find_element_by_xpath("//*[@id='flatGroup']/li[3]/div/div[4]/img")
        #tickButton.click()
        #click Search Button
        time.sleep(3)

        self.wait_for_class_to_be_available(self.browser,self.searchButton)

        self.wait_for_class_to_be_available(self.browser,self.selectAllCheck)
        
        time.sleep(5)

        givenTotalPieces = self.wait_for_xpath_to_present(self.browser,self.selectedPieces)
        givenTotalCtn = self.wait_for_xpath_to_present(self.browser,self.selectedCtn)
        givenNetWeight = self.wait_for_xpath_to_present(self.browser,self.selectedNetWeight)
        givenGrossWeight = self.wait_for_xpath_to_present(self.browser,self.selectedGrossWeight)
        givenTotalVolume = self.wait_for_xpath_to_present(self.browser,self.selectedTotalVolume)


        #print(self.browser.find_element_by_xpath("//*[@id='selectedDataLeft']/div/table/tbody/tr/td[3]/ul/li/label").get_attribute('innerHTML'))

        #lbl = li.find_element_by_tag_name('label')

        #print(lbl.text)
        givenTotalPieces = "{:.2f}".format(float(givenTotalPieces))
        givenNetWeight = "{:.1f}".format(float(givenNetWeight))
        givenGrossWeight = "{:.1f}".format(float(givenGrossWeight))
        givenTotalVolume = "{:.1f}".format(float(givenTotalVolume))

        print(givenTotalPieces)
        print(givenTotalCtn)
        print(givenNetWeight)
        print(givenGrossWeight)
        print(givenTotalVolume)

        print(self.totalPieces)
        print(self.totalCtn)
        print(self.netWeight)
        print(self.grossWeight)
        print(self.totalVolume)

        print(type(givenTotalCtn))
        print(type(self.totalCtn))

        if (givenTotalPieces.strip() == self.totalPieces.strip()) and (givenTotalCtn.strip() == self.totalCtn.strip()) and (givenNetWeight.strip() == self.netWeight.strip()) and (givenGrossWeight.strip() == self.grossWeight.strip()) and (givenTotalVolume.strip() == self.totalVolume.strip()):
            print('GG WP EZ PZ')

            self.wait_for_class_to_be_available(self.browser,self.tickAtBottomMenu)

            self.wait_for_class_to_be_available(self.browser,self.consolidationDiv)

            self.wait_for_class_to_be_available(self.browser,self.filterButton)

            enterInput = self.browser.find_element_by_xpath(self.poEntryInput)
            enterInput.send_keys(self.orderNumbers)


            self.wait_for_class_to_be_available(self.browser,self.tickButton)
        
            self.wait_for_class_to_be_available(self.browser,self.searchButton)

            self.wait_for_class_to_be_available(self.browser,self.selectAllCheck)

            time.sleep(5)

            givenTotalPieces = self.wait_for_xpath_to_present(self.browser,self.totalPiecesXPath)
            givenTotalCtn = self.wait_for_xpath_to_present(self.browser,self.totalCtnXPath)
            givenNetWeight = self.wait_for_xpath_to_present(self.browser,self.netWeightXPath)
            givenGrossWeight = self.wait_for_xpath_to_present(self.browser,self.grossWeightXPath)
            givenTotalVolume = self.wait_for_xpath_to_present(self.browser,self.totalVolumeXPath)

            selectedTotalPieces = self.wait_for_xpath_to_present(self.browser,self.selectedPieces)
            selectedTotalCtn = self.wait_for_xpath_to_present(self.browser,self.selectedCtn)
            selectedNetWeight = self.wait_for_xpath_to_present(self.browser,self.selectedNetWeight)
            selectedGrossWeight = self.wait_for_xpath_to_present(self.browser,self.selectedGrossWeight)
            selectedTotalVolume = self.wait_for_xpath_to_present(self.browser,self.selectedTotalVolume)

            if (givenTotalPieces.strip() == selectedTotalPieces.strip() and givenTotalCtn.strip() == selectedTotalCtn.strip()):
                
                print('Final GG WP')
                #Clicking Container

                self.wait_for_class_to_be_available(self.browser,self.containerButton)
                #clicking create container button

                self.wait_for_class_to_be_available(self.browser, self.createContainer)

                elementFind = WebDriverWait(self.browser, 100).until(EC.presence_of_element_located((By.NAME, 'treTransportMode')))
                element = self.browser.find_element_by_xpath("//select[@name='treTransportMode']")
                all_options = element.find_elements_by_tag_name("option")
                for option in all_options:
                    if (option.get_attribute("value") == 'SEA'):
                        option.click()
                

                container_textBox = WebDriverWait(self.browser, 100).until(EC.presence_of_element_located((By.NAME, 'treContainorNumber')))
                container_textBox.clear()
                container_textBox.send_keys(self.containerNumber)

                seal_textBox = WebDriverWait(self.browser, 100).until(EC.presence_of_element_located((By.NAME, 'trePlombNumber')))
                seal_textBox.send_keys(self.sealNumber)

                time.sleep(1)
                containerType = self.browser.find_element_by_xpath("//select[@name='eqpId']")
                all_options = containerType.find_elements_by_tag_name("option")
                for option in all_options:
                    if (option.get_attribute("value") == '6'):
                        option.click()
                time.sleep(2)
                checkBoxItem = WebDriverWait(self.browser, 100).until(EC.element_to_be_clickable((By.NAME, 'treIdSelected')))
                
                checkBoxItem.click()

                self.wait_for_class_to_be_available(self.browser,"//*[@title='button.addAs']")
                
                print('Heavy')

                

                '''time.sleep(1)
                routerElement = self.browser.find_element_by_xpath("//select[@name='lneId']")
                all_options = routerElement.find_elements_by_tag_name("option")
                for option in all_options:
                    if (option.get_attribute("value") == 'ASIA_PLN5068'):
                        option.click()'''


        #4
        #self.bar.next()

    def iterateOverInput(self):
        
        #print(self.inputs)

        self.orderNumbers = self.inputs['ORDER NUMBER'].values.tolist()
        print(self.orderNumbers)
        
        self.orderNumbers = ' '.join(str(e) for e in self.orderNumbers)
        
        self.totalPieces = self.inputs['PIECES PER PO'].sum()
        self.totalPieces = "{:.2f}".format(self.totalPieces)

        self.netWeight = self.inputs['NET WEIGHT'].sum()
        self.netWeight = "{:.1f}".format(self.netWeight)

        self.totalCtn = str(self.inputs['CARTONS'].sum())

        self.grossWeight = self.inputs['GROSS WEIGTH'].sum()
        self.grossWeight = "{:.1f}".format(self.grossWeight)

        self.totalVolume = self.inputs['VOLUME'].sum()
        self.totalVolume = "{:.1f}".format(self.totalVolume)
        
        self.containerNumber = self.inputs['CONTAINER NUMBER'][0]
        self.containerSize = self.inputs['CONTAINER SIZE'][0]
        self.trType = self.inputs['TR TYPE'][0]
        self.sealNumber = self.inputs['SEAL NUMBER'][0]

        print(self.totalPieces)
        print(self.netWeight)
        print(str(self.totalCtn))
        print(self.grossWeight)
        print(self.totalVolume)

        print('///////////////////')

        print(self.inputs['CONTAINER NUMBER'][0])
        print(self.inputs['SEAL NUMBER'][0])
        print(self.inputs['TR TYPE'][0])
        print(self.inputs['CONTAINER SIZE'][0])
        pathForDropdown = "//select[@id='treTransportModeASIA_CNT1283764']/option[contains(text(),'"+  self.trType +"')]"
        print(pathForDropdown)
        


if __name__ == '__main__':
    obj = downloader()
    obj.iterateOverInput()
    obj.loadBrowser()