__author__ = 'nbf1707'
#python 2.7.6

from bs4 import BeautifulSoup   #or use import bs4, then bs4.BeautifulSoup later
from splinter import Browser
successFlag = True

def Login(userName, Password):
    br = Browser('chrome')
    urlLogin = "https://bthealth.service-now.com/welcome.do"
    br.visit(urlLogin)
    soup = BeautifulSoup(br.html)
    if soup.get_text().find("Login") == -1:     #can't find Login page
        br.quit()
        return br, False
    #print br.html
    br.fill("user_name", userName)
    br.fill("user_password", Password)
    button = br.find_by_name('not_important')
    button.click()
    soup = BeautifulSoup(br.html)
    if soup.get_text().find("Welcome:") == -1:  #didn't login ok
        br.quit()
        return br, False
    return br, True

def logSR(br, urlOfSR, LocalRef, Comment, submit=False, attachPath=""):
    br.visit(urlOfSR)
    br.fill("IO:1db28ec14a36232800446e9bf722ab6c", LocalRef)
    br.fill("IO:3823341e0a0a0b27003118193596953f", Comment)
    br.check("ni.IO:356262034a36232800d671ef05504e25")  #Terms and Conditions
    if attachPath != "":    #attachment
        br.execute_script("saveCatAttachment(gel('sysparm_item_guid').value, 'sc_cart_item')")  #add attachment
        br.attach_file("attachFile", attachPath)
        button = br.find_by_id('attachButton')
        button.click()
        closeButton = br.find_by_id('popup_close_image')
        closeButton.click()
    button = br.find_by_id('order_now')
    button.click()
    if submit:
        #br.execute_script("forms['service_catalog.do'].submit()")  #submit order
        print "I've ordered this!"
    return br, True

def getRITMNumber(br):
    try:
        soup = BeautifulSoup(br.html)
        if soup.get_text().find("RITM") == -1:     #can't find number
            return '', False
        else:
            return soup.find("td", {"class" : "checkoutNumber"}).text, True
    except:
        return '', False