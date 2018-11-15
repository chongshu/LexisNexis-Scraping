# -*- coding: utf-8 -*-
"""
Created on Mon Nov 12 17:31:39 2018

@author: chongshu
"""


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

import re
import numpy as np
import os
import time
import zipfile
from docx import Document
import pandas as pd
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from xml.etree.cElementTree import XML
import sys
  
searchTerms = r'apple shareholder class action'
url = r'http://libguides.usc.edu/go.php?c=9232127'
username = 'MyUSCPassUsername'
password = 'MyUSCPassWord'
root = r'C:\Users\chongshu\Desktop\LexisNexis'
path_to_chromedriver = root + r'\chromedriver'
download_folder = root + r'\download'
dead_time = 300


def download_file(url = url, searchTerms = searchTerms, username = username, \
                  dead_time = dead_time, path_to_chromedriver=path_to_chromedriver, \
                  download_folder = download_folder):
    while True:
        try:
            chromeOptions = webdriver.ChromeOptions()
            prefs = {"download.default_directory" : download_folder}
            chromeOptions.add_experimental_option("prefs",prefs)
            browser = webdriver.Chrome(executable_path = path_to_chromedriver, options=chromeOptions)
            browser.set_window_size(1800, 1000)
            #Login
            browser.get(url)
            browser.find_element_by_id('username').send_keys(username)
            browser.find_element_by_id('password').send_keys(password)
            browser.find_element_by_xpath('//*[@id="loginform"]/div[4]/button').click()
            # Get Page Info
            browser.find_element_by_xpath('//*[@id="searchTerms"]').send_keys(searchTerms)
            browser.find_element_by_xpath('//*[@id="mainSearch"]').click()
            N_temp = browser.find_element_by_xpath('//*[@id="content"]/header/h2/span').text
            time.sleep(5)
            total_number = int(''.join(re.findall(r'[0-9]', N_temp)))
            total_page = int(np.ceil(total_number/10))
            file_digit = len(str(total_page)) * 2 + 1
            # Sort by Date
            start_time = time.time()
            while True:
                if time.time() - start_time > dead_time:
                    raise Exception()
                try:        
                    browser.find_element_by_xpath('//*[@id="results-list-delivery-toolbar"]/div/ul[2]/li/div/button').click()
                    break
                except WebDriverException: 
                    pass             
            browser.find_element_by_xpath('//*[@id="results-list-delivery-toolbar"]/div/ul[2]/li/div/div/button[4]').click()
            for page in range(1, total_page + 1):
# =============================================================================
#     If the file already exists. Go to next page. 
# =============================================================================
                if os.path.isfile(download_folder + '\\'  + (str(page) + '_' + str(total_page)).zfill(file_digit) +  '.ZIP'):
                    print('exist: ' + str(page) + '_' + str(total_page) +  '.ZIP')
                    if page < total_page:
                        try: 
                            start_time = time.time()
                            while True:
                                if time.time() - start_time > dead_time:
                                    raise Exception()
                                WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.LINK_TEXT , str(page + 1))))
                                browser.find_element_by_link_text(str(page + 1)).click()
                        except WebDriverException: 
                            time.sleep(3)
                            continue         
                    else:
                        print('FINISHED!')
                    
                        
# =============================================================================
# Wait the all checkbox to be clickable
# =============================================================================
                start_time = time.time()
                while True:
                    if time.time() - start_time > dead_time:
                        raise Exception()
                    try:
                        if browser.find_element_by_xpath('//*[@id="results-list-delivery-toolbar"]/div/ul[1]/li[1]/input').get_attribute('checked') != 'true':
                            browser.find_element_by_xpath('//*[@id="results-list-delivery-toolbar"]/div/ul[1]/li[1]/input').click()
                        break
                    except WebDriverException: 
                        pass
                   
# =============================================================================
#   Wait the IncludeAttachments button to be clickable.  otherwise re-click download buttom 
# =============================================================================
                time.sleep(1)
                try: 
                    start_time = time.time()
                    while True:
                        if time.time() - start_time > dead_time:
                            raise Exception()
                        elm = browser.find_element_by_xpath('//*[@id="results-list-delivery-toolbar"]/div/ul[1]/li[4]/ul/li[3]/button')
                        elm.click()
                except WebDriverException: 
                    pass     
                start_time = time.time()
                while True:
                    if time.time() - start_time > dead_time:
                        raise Exception()

                    try:
                        browser.find_element_by_xpath('//*[@id="DocumentsOnly"]').click()
                        browser.find_element_by_xpath('//*[@id="IncludeAttachments"]').click()
                        break
                    except WebDriverException: 
                        pass
                browser.find_element_by_xpath('//*[@id="Docx"]').click()
                browser.find_element_by_xpath('//*[@id="SeparateFiles"]').click()
                browser.find_element_by_xpath('//*[@id="FileName"]').clear()
                browser.find_element_by_xpath('//*[@id="FileName"]').send_keys((str(page) + '_' + str(total_page)).zfill(file_digit))      
# =============================================================================
#     After downloading Close the pop up window. LexisNexis only allows 5 cocurrent windows
# =============================================================================
                before = browser.window_handles[0]
                browser.find_element_by_xpath('/html/body/aside/footer/ul/li[1]/input').click()
                start_time = time.time()
                while True:
                    if time.time() - start_time > dead_time:
                        raise Exception()
                    try:
                        after = browser.window_handles[1]
                        break
                    except: 
                        pass
                browser.switch_to.window(after)
                start_time = time.time()
                while True:
                    if time.time() - start_time > dead_time:
                        raise Exception()
                    if  browser.find_elements_by_link_text((str(page) + '_' + str(total_page)).zfill(file_digit)):
                        break
                browser.close()
                browser.switch_to.window(before)   
                print('finishing page' + str(page) + '_' + str(total_page)  ) 
# =============================================================================
#     Go to the next page. Have to check the next page link is clickable
# =============================================================================
                if page < total_page:
                    try: 
                        browser.find_element_by_link_text(str(page + 1)).click()
                        start_time = time.time()
                        while True:
                            if time.time() - start_time > dead_time:
                                raise Exception()
                            browser.find_element_by_link_text(str(page + 1)).click()
                    except WebDriverException: 
                        time.sleep(3)
                        pass
                else:
                    print('FINISHED!')
                    sys.exit()
        except Exception:
            print('restarting')
            continue
                        

def unzip(download_folder=download_folder):
    if not os.path.exists(download_folder + '\\' + 'unzipped'):
        os.makedirs(download_folder + '\\' + 'unzipped') 
    for filename in os.listdir(download_folder):
        if not filename.endswith('.ZIP'):
            continue
        zip_ref = zipfile.ZipFile(download_folder +  '\\' + filename, 'r')
        zip_ref.extractall(download_folder + '\\' + 'unzipped')
        if len(zip_ref.namelist()) != 11:
            print('missing document at ' + filename)
        zip_ref.close()
        print('unzipping ' + filename)
        

def create_index(download_folder=download_folder, searchTerms = searchTerms):  
    def get_docx_text(path):
        WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        PARA = WORD_NAMESPACE + 'p'
        TEXT = WORD_NAMESPACE + 't'
        """
        Take the path of a docx file as argument, return the text in unicode.
        """
        document = zipfile.ZipFile(path)
        xml_content = document.read('word/document.xml')
        document.close()
        tree = XML(xml_content)
    
        paragraphs = []
        for paragraph in tree.getiterator(PARA):
            texts = [node.text
                     for node in paragraph.getiterator(TEXT)
                     if node.text]
            if texts:
                paragraphs.append(''.join(texts))
    
        return '\n\n'.join(paragraphs)
    def get_docx_hyperlink(path):
        document = Document(path)
        rels = document.part.rels
        link = []
        for rel in rels:
            if rels[rel].reltype == RT.HYPERLINK:
                link.append(rels[rel]._target)
        return pd.Series(link)   
    def find_between(s, first, last):
        list = []
        try:
            while True:
                try:
                    start = s.index( first ) + len( first )
                    end = s.index( last, start )
                    list.append( s[start:end])
                    s = s.replace(first + s[start:end] + last,'')
                except:
                    break
            return list
        except ValueError:
            return ""
        
    index = pd.DataFrame(({'Title':{},'Link':{}}))
    for filename in os.listdir(download_folder + '\\' + 'unzipped'):
        if filename.find("doclist") == -1:
            continue
        text = get_docx_text(download_folder + '\\' + 'unzipped\\' + filename)
        link = get_docx_hyperlink(download_folder + '\\' + 'unzipped\\' + filename).loc[0]
        content = '<start>' + re.sub(r'.*Documents \(\d*\)', '',text.replace('\n',' ')).replace(\
              'Client/Matter: -None-  Search Terms: '+ searchTerms + '  Search Type: Terms and Connectors   Narrowed by:   Content Type  Narrowed by  News  -None-',\
              '<end><start>')
        content_list = pd.Series(find_between(content, '<start>', '<end>'))
        content_list = content_list.str[5:]
        index = index.append(pd.DataFrame({'Title':content_list,'Link':link}),ignore_index = True)
    index = index.reset_index()
    index['index'] = index.index + 1
    index.to_csv(download_folder + '\\index.csv', index=False)
    return index





