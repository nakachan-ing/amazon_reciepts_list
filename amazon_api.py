#!/usr/bin/env python
# coding: utf-8

# In[35]:


from bs4 import BeautifulSoup
import os
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import json


# In[28]:


term = input('取得したい期間を入力してください( 過去30日間/過去3か月/2015~2022): ') + "年"

url = 'https://www.amazon.co.jp/'

# Chromeの拡張機能でPDF保存するための設定
downloadPath ='/Users/nakadakyota/python/autorun/amazon_reciepts_list'
options = webdriver.ChromeOptions()
settings = {"recentDestinations": [{"id": "Save as PDF",
                                    "origin": "local",
                                    "account": ""}],
            "selectedDestinationId": "Save as PDF",
            "version": 2}
prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),'savefile.default_directory':downloadPath}
options.add_experimental_option('prefs', prefs)
options.add_argument('--kiosk-printing')

# ChromeのPDF保存オプションを指定してブラウザを起動させる
browser = webdriver.Chrome(options=options)
browser.implicitly_wait(3)
browser.get(url)


# In[29]:


login = browser.find_element_by_xpath('//*[@id="nav-link-accountList"]')
login.click()


# In[30]:


# ログインIDとパスワードを環境変数として設定
# os.environ["MY_LOGIN_ID"] = "YOUR_LOGIN_ID"
# os.environ["MY_LOGIN_PASSWORD"] = "YOUR_LOGIN_PASSWORD"

email = os.environ.get("AMAZON_LOGIN_ADDRESS")
passwd = os.environ.get("AMAZON_LOGIN_PASS")

email_element = browser.find_element_by_xpath('//*[@id="ap_email"]')
email_element.clear()
time.sleep(1)
email_element.send_keys(email)
next_page = browser.find_element_by_xpath('//*[@id="continue"]')
time.sleep(1)
next_page.click()


# In[31]:


passwd_element = browser.find_element_by_xpath('//*[@id="ap_password"]')
passwd_element.clear()
time.sleep(1)
passwd_element.send_keys(passwd)
submit = browser.find_element_by_xpath('//*[@id="signInSubmit"]')
time.sleep(1)
submit.click()


# In[32]:


order_history = browser.find_element_by_xpath('//*[@id="nav-orders"]')
time.sleep(1)
order_history.click()


# In[34]:


# 領収書をダウンロードし、次のページに遷移するを繰り返してページ分実行

selection = browser.find_element_by_xpath('//*[@id="a-autoid-1-announce"]')
selection.click()
choose = browser.find_element_by_link_text(term)
choose.click()

while True:
    main_brows = browser.current_window_handle
    reciept_links = browser.find_elements_by_link_text('領収書等')

    for reciept_link in reciept_links:
        reciept_link.click()
        time.sleep(3)
        reciept_purcahseresult_link = browser.find_element_by_link_text('領収書／購入明細書')

        before_click_handles = browser.window_handles

        
        actions = ActionChains(browser)
        actions.key_down(Keys.COMMAND)
        actions.click(reciept_purcahseresult_link)
        actions.perform()

        WebDriverWait(browser, 5).until(lambda a: len(browser.window_handles) > len(before_click_handles))

        after_click_handles = browser.window_handles
        handle_new = list(set(after_click_handles) - set(before_click_handles))
        
        browser.switch_to.window(handle_new[0])

        
        browser.execute_script('window.print();')

        browser.close()
        browser.switch_to.window(main_brows)

    try:
        browser.find_element_by_class_name('a-last').find_element_by_tag_name('a')
        browser.find_element_by_class_name('a-last').click()
        browser.switch_to.window(browser.window_handles[-1])
    except:
        break

print('領収書の保存が完了しました')


# In[ ]:




