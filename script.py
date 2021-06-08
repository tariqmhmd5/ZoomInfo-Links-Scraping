from selenium import webdriver
import time
import numpy as np
from bs4 import BeautifulSoup
import pandas as pd
import pymsgbox
import easygui
from pandas import ExcelWriter
from pandas import ExcelFile
import re

wd = webdriver.Chrome()
wd.maximize_window()
wd.get('https://chrome.google.com/webstore/detail/bulk-url-opener/pknlccgephlicnahemjfcfnbifiglegd')
easygui.msgbox("Add this Chrome Extension after that press Ok to Continue", title="Alert!")
time.sleep(4)
#wd.find_element_by_xpath('/html/body/div[3]/div[2]/div/div/div[2]/div[2]/div/div/div/div').click()
#pymsgbox.alert('Click on Add app in Chhrome after that press Ok to Continue', 'Alert!')

try:
    wd.switch_to.window(wd.window_handles[1])
    wd.close()
    wd.switch_to.window(wd.window_handles[0])
    wd.close()
    wd.switch_to.window(wd.window_handles[0])
    wd.get('chrome-extension://pknlccgephlicnahemjfcfnbifiglegd/index.html')
    n = pymsgbox.prompt('Add and open urls in chrome. How Many Urls you have pasted?')
    n = int(n)
except Exception as e:
    print('Restart the app')
    print(e)

try:
    for i in range(1,n+1):
        wd.switch_to.window(wd.window_handles[i])
        links = []
        source = wd.page_source
    
        soup = BeautifulSoup(source)
        for link in soup.findAll('a'):
            links.append(link.get('href'))
    
        output = pd.DataFrame(links,columns=['links']) 
        writer = ExcelWriter('url_'+str(i)+'.xlsx')
        output.to_excel(writer,'Sheet1')
        writer.save()
    pymsgbox.alert('Task Complete!', 'Alert!')
except Exception as e:
    print('Restart the app, an error occured')
    print(e)
