import os
import time

import argparse

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

def urlToName(url):
    name =  url.split('/')[-1]
    return name

def interval(value):
    tInterval={
        '1G':'a[1]',
        '1S':'a[2]',
        '1M':'a[3]',
        '3M':'a[4]',
        '6M':'a[5]',
        '1A':'a[6]',
        '5A':'a[7]',
        '10A':'a[8]'
        }
    if value not in tInterval:
        print('No such value, return default')
        return tInterval.get('1G')
    return tInterval.get(value)

def downloadData( stockName, downloadUrl, stockInterval, downloadFolder=None ) :
    fileName = "amCharts.csv" #name of file that the browser normally downloads

    newFolder = args.dFolder + os.sep + '.tmp_' + stockName 
    if not os.path.isdir(downloadFolder):
        os.mkdir(newFolder,0x666) # makes temporary folder
        print("Temporary folder created")

    #while cycle needed due to refresh behaviour of the target page

    start = time.time()
    
    options=Options()
    options.headless = True
    
    profile = webdriver.firefox.firefox_profile.FirefoxProfile()
    profile.set_preference('browser.download.folderList',2)
    profile.set_preference('browser.download.dir',newFolder)
    profile.set_preference("browser.download.useDownloadDir",True)
    profile.set_preference('browser.download.manager.showWhenStarting',False)
    profile.set_preference("browser.helperApps.alwaysAsk.force",False)
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk',(\
                            'text/plain,\
                            text/csv,\
                            application/csv,\
                            application/msexcel,\
                            application/octet-stream,\
                            text/javascript,\
                            application/vnd.ms-excel,\
                            application/vnd.openxmlformats\
                            -officedocument.spreadsheetml.sheet'))
    
    driver = webdriver.Firefox(firefox_profile=profile,options=options)

    agreeButton = '//*[@id="onetrust-accept-btn-handler"]'
    baseXpath=  '/html/body/div/div[1]/div/div[1]/div[4]/'\
                'div[1]/section[2]/div/div[1]' #xpath common to all elements
    if '.Q' in stockName:
        baseXpath = baseXpath + '/div[2]/div/div/div/div[1]'
    else:
        baseXpath = baseXpath + '/div/div[2]/div[1]/div[2]/div/div/div[1]'

    downIco =   baseXpath + '/div[2]/div/div[1]/div[2]/ul/li'

    optionMenu= downIco + "/ul" #xpath to sub-menu OK

    csvFile = optionMenu + "/li[3]/a" #xpath of CSV option OK
    
    while( not os.path.isfile( newFolder+os.sep+fileName )):         
        try:
            wait = WebDriverWait(driver,10)
            driver.get(downloadUrl) 
            
            pointer = wait.until(\
                    EC.presence_of_element_located((By.ID,"onetrust-policy")))

            pointer = driver.find_element_by_xpath(agreeButton)
            pointer.click()
            print("cookies accepted")
            
            driver.execute_script("window.scrollTo(0,500);")
            print("Page scrolled")
            
            if stockInterval != '1G':
                timeInterval =  baseXpath + \
                                '/div[1]/div[2]/' + \
                                interval(stockInterval)
                pointer = wait.until(\
                    EC.visibility_of_element_located((By.XPATH,timeInterval)))
                pointer = driver.find_element_by_xpath(timeInterval)
                pointer.click()
                time.sleep(4) # wait for chart to refresh

            pointer = driver.find_element_by_xpath(downIco)
            print("Download icon found")
            hover = ActionChains(driver).move_to_element(pointer)
        
            pointer = driver.find_element_by_xpath(csvFile)
            hover.move_to_element(pointer) #Moves to CSV option
            hover.click()    
            hover.perform()
            print("Download actions performed\nStarting Download...")

            time.sleep(1) #Ensures full download of the data
        
            driver.close()
        except NoSuchElementException:
                print("Nothing found")       
                driver.close()

    end = time.time()
    print("Duration of download: {}".format(end-start))
        
    print("File correctly downloaded, starting renaming...")
         
    stockDate = time.strftime("%d-%m-%Y")  #Renaming the fresh-downloaded file
    
    os.rename(  newFolder + os.sep + fileName, \
                downloadFolder + os.sep + \
                stockName + '_' + stockInterval + \
                '_' + stockDate + '.csv')

    if os.path.isdir(newFolder):
        os.rmdir(newFolder)
        print("Temporary folder deleted")

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Automates the download of '\
                                    '.csv files from ilsole24ore.com website')
    parser.add_argument(
            '-d', \
            help='download/path where the new .csv file will be stored', \
            dest='dFolder', \
            default= os.path.expanduser('~')+os.sep+"Downloads")
    parser.add_argument(
            '-u', \
            help=   'Url from which to download the data\n' \
                    'Default set to "https://mercati.ilsole24ore.com/"\
                    "materie-prime/commodities/oro/GLDEUR"', \
            dest = 'dUrl', \
            default =   "https://mercati.ilsole24ore.com/materie-prime"\
                        "/commodities/oro/GLDEUR"
            )
    parser.add_argument(
            '-i', \
            help=   'Time interval of the data in chart', \
            default = '1G', \
            choices = ['1G','1S','1M','3M','6M','1A','5A','10A'], \
            dest = 'tInterval'
            )
    parser.add_argument(
            '-rf', \
            help='Reference file from which the download url will be drawn', \
            dest = 'refFile'
            ) 
    
    args = parser.parse_args()

    if args.refFile != None :
        
        threads = []
        for url in open(args.refFile,'r'):            
            url = url.replace('\n','') 
            stockName = urlToName(url)        
            downloadData( urlToName(url),url,\
                          args.tInterval,args.dFolder)

    else:
      downloadData(urlToName(args.dUrl),args.dUrl,args.tInterval,args.dFolder)
    
    print("Total threads (1 master): {}".format(threading.activeCount()))

