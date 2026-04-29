# -*- coding: utf-8 -*-
"""
Created on Tue Oct  9 16:07:18 2018

@author Jonathan Estrada

"""

from selenium import webdriver
import xlrd
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from base64 import b64encode
import random
from selenium.webdriver.common.proxy import Proxy,ProxyType,ProxyTypeFactory
import shutil
import os
import stat

DEBUG = False



def read_xlsx(inputfile):
    workbook = xlrd.open_workbook(inputfile)
    worksheet = workbook.sheet_by_index(0)
    first_row = [] # The row where we stock the name of the column
    for col in range(worksheet.ncols):
        first_row.append( worksheet.cell_value(0,col) )
    # transform the workbook to a list of dictionaries
    data =[]
    for row in range(1, worksheet.nrows):
        elm = {}
        for col in range(worksheet.ncols):
            elm[first_row[col]]=worksheet.cell_value(row,col)
        data.append(elm)
    return first_row,data



def loginGmail(B,login,password,remail):
    while True:
        try:
            B.get("https://accounts.google.com")
            break
        except:
            time.sleep(3)
    B.find_element_by_id("identifierId").send_keys(login)
    B.find_element_by_id("identifierNext").click()
    time.sleep(2)
    while True:
        try:
            B.find_element_by_name("password").send_keys(password)
            break
        except:
            time.sleep(1)
    B.find_element_by_id("passwordNext").click()
    time.sleep(3)
    if "signin/v2/challenge" in B.current_url:
        B.find_element_by_class_name("C5uAFc").click()
        time.sleep(1)
        while True:
            try:
                B.find_element_by_id("identifierId").send_keys(remail)
                break
            except:
                time.sleep(1)
        for i in B.find_elements_by_class_name("U26fgb"):
            if "Next" in i.text or "Sus" in i.text:
                i.click()
                break
        time.sleep(3)
    time.sleep(1)
    print "Logged"


def answer(B,k):
    for i in B.find_elements_by_class_name("section-verify-edit-"+k):
        if i.is_displayed():
            i.click()
            break

def answer2(B,k):
    ie = 0
    for i in B.find_elements_by_class_name("section-verify-edit-"+k):
        if i.is_displayed():
            if ie==0:
                ie+=1
            else:
                i.click()
                break


def goDown(B):
    for i in range(5):
        try:
            html = B.find_elements_by_class_name('section-listbox')[1]
            break
        except:
            time.sleep(1)
    html.send_keys(Keys.PAGE_DOWN)
    time.sleep(0.8)

def answers(B,question):
    if "closed" in question or "hour" in question or "Type" in question or "Status" in question or "Business Hours" in question:
        print "YES"
        answer(B,"yes")
    elif  "Category" in question or "Spam" in question or "Website" in question or "category" in question or "spam" in question or "website" in question:
        print "NOT SURE"
        answer(B,"not-sure")
    elif "phone" in question or "Phone" in question:
        print "NO"
        answer(B,"no")
    else:
        print "NOT SURE"
        answer(B,"not-sure")
    time.sleep(1)

def answers2(B,question):
    if "closed" in question or "hour" in question or "Type" in question or "Status" in question or "Business Hours" in question:
        print "YES"
        answer2(B,"yes")
    elif  "Category" in question or "Spam" in question or "Website" in question or "category" in question or "spam" in question or "website" in question:
        print "NOT SURE"
        answer2(B,"not-sure")
    elif "phone" in question or "Phone" in question:
        print "NO"
        answer2(B,"no")
    else:
        print "NOT SURE"
        answer2(B,"not-sure")
    time.sleep(1)

def zoom_out(B):
    #B.find_element_by_id("widget-zoom-out").click()
    B.find_element_by_id("widget-zoom-out").click()
    time.sleep(3)
    for i in range(5):
        try:
            B.find_element_by_class_name("widget-search-this-area-inner").click()
            break
        except:
            time.sleep(1)
    time.sleep(6)
    
 
def clickok(B):
    try:
        if B.find_element_by_id("last-focusable-in-modal").is_displayed():
            B.find_element_by_id("last-focusable-in-modal").click()
            return True
    except:
        pass
    
    
def questionAnswers(B,repeat_times,randomwait):
    for i in range(6):
        print i," TIMES"
        iId = ""
        while repeat_times:
            print repeat_times
            if repeat_times%3==0:
                goDown(B)
            
            
            if len(B.find_elements_by_class_name("section-verify-edit-title"))==0:
                print "There are not exist edits"
                break
            # Get question
            question = None
            while question is None:
                for i in B.find_elements_by_class_name("section-verify-edit-title"):
                    if i.is_displayed():
                        question = i.text
                        break
                else:
                    time.sleep(2)
            if iId == i.id:
                print "SAME ID"
                if not clickok(B):
                    print "NO BLOCK BREAK"
                    break
            else:
                iId = i.id
            # Answer on question        
            answers(B,question)            
            # Done
            repeat_times=repeat_times-1
            if repeat_times<=0:
                return 0
            time.sleep(0.3)
            


            # Answer on second question if exist
            if "Verify 1 more" in B.find_element_by_tag_name("body").text:
                question = None
                ik = 0
                while question is None:
                    for i in B.find_elements_by_class_name("section-verify-edit-title"):
                        if i.is_displayed():
                            if ik==0:
                                ik+=1
                            else:
                                question = i.text
                                break
                    else:
                        time.sleep(1)                            
                print "2 = ",question
                answers2(B,question)
                repeat_times-=1
                if repeat_times==0:
                    return 0
                time.sleep(1)
            time.sleep(random.randint(int(randomwait[0]),int(randomwait[-1])))
        # DO ZOOM OUT 2 times
        time.sleep(3)
        print "ZOOM OUT"
        zoom_out(B)
    return repeat_times

def openGoogleMap(B,city,repeat_times):
    
    while True:
        try:
            B.get("https://maps.google.com/localguides/home/onboarding")
            break
        except:
            time.sleep(2)
            print "There is internet or chrome problem look at it! Trying again to oprn google local guides"
    time.sleep(3)
    try:
        for i in B.find_elements_by_tag_name("button"):
            if "SKIP" in i.text:
                i.click()
                time.sleep(4)
                break
    except:
        pass
    try:
        elem = B.find_element_by_name("selectedCity")
        elem.clear()
        elem.send_keys(city)
        time.sleep(2)
        im = B.find_elements_by_class_name("md-autocomplete-suggestions-container")[-1]
        B.execute_script("arguments[0].className  = 'md-autocomplete-suggestions-container md-whiteframe-z1 md-virtual-repeat-container md-orient-vertical';", im) 
        time.sleep(2)
        try:
            B.find_elements_by_tag_name("ul")[-1].find_element_by_tag_name("li").click()
        except:
            B.find_element_by_class_name("ng-scope").click()
        time.sleep(1)
        B.execute_script("arguments[0].className  = 'md-autocomplete-suggestions-container md-whiteframe-z1 md-virtual-repeat-container md-orient-vertical ng-hide';", im) 
        time.sleep(2)
        B.find_element_by_name("ageConsent").click()
        B.find_element_by_name("emailConsent").click()
        time.sleep(2)
        for i in B.find_elements_by_tag_name("button"):
            if "SIGN UP" in i.text:
                i.click()
                time.sleep(3)
                break
    except:
        pass
    
    try:
        for i in B.find_elements_by_tag_name("button"):
            if "SKIP" in i.text:
                i.click()
                time.sleep(4)
                break
    except:
        pass
    
    while True:
        try:
            B.get("https://maps.google.com/localguides/home/settings")
            B.get("https://maps.google.com/localguides")
            break
        except:
            time.sleep(3)
            print "There is internet or chrome problem look at it! Trying again to oprn google localguides"
        
    time.sleep(3)
    try:
        ccount = int(B.find_element_by_tag_name("lg-animated-number").text.split()[0].replace(",",""))
        repeat_times = abs(repeat_times-ccount)
    except:
        pass
    try:
        B.find_element_by_class_name("md-raised").click()
        time.sleep(1)
        elem = B.find_element_by_id("input-1")
        elem.send_keys(city)
        time.sleep(2)
        im = B.find_element_by_class_name("md-autocomplete-suggestions-container")
        B.execute_script("arguments[0].className  = 'md-autocomplete-suggestions-container md-whiteframe-z1 md-virtual-repeat-container md-orient-vertical';", im) 
        time.sleep(2)
        try:
            B.find_element_by_id("ul-1").find_element_by_tag_name("li").click()
        except:
            B.find_element_by_class_name("ng-scope").click()
        time.sleep(1)
        B.execute_script("arguments[0].className  = 'md-autocomplete-suggestions-container md-whiteframe-z1 md-virtual-repeat-container md-orient-vertical ng-hide';", im) 
        time.sleep(2)
        B.find_element_by_name("ageConsent").click()
        B.find_element_by_name("emailConsent").click()
        B.find_element_by_class_name("md-raised").click()
        time.sleep(2)
        for i in B.find_elements_by_tag_name("button"):
            if "SIGN UP" in i.text:
                i.click()
                time.sleep(3)
                break
    except:
        pass



    time.sleep(3)
    try:
        for i in B.find_elements_by_tag_name("button"):
            if "SKIP" in i.text:
                i.click()
                time.sleep(4)
                break
    except:
        pass
    
    print "Search city"
    B.get('https://www.google.com/maps')
    B.find_element_by_id("searchboxinput").send_keys(city)
    B.find_element_by_id("searchboxinput").send_keys(Keys.RETURN)
    time.sleep(5)
    print "Open contribution"
    B.find_element_by_class_name("searchbox-hamburger").click()
    rdc = 0
    while True:
        if rdc == 5:
            try:
                B.find_element_by_class_name("searchbox-hamburger").click()
            except:
                pass
            rdc = 0
        try:
            time.sleep(1)
            B.find_element_by_class_name("maps-sprite-settings-rate-review").click()
            time.sleep(1)
            break
        except:
            rdc+=1
    dd = 0
    while True:
        if dd == 100:
            break
        dd+=1
        try:
            B.find_elements_by_class_name("section-list-item-content")[1].click()
            break
        except:
            time.sleep(1)
    time.sleep(4)
    try:
        B.find_element_by_id("last-focusable-in-modal").click()
    except:
        pass
    time.sleep(1)



    return repeat_times
        
        
fp=None


def copytree(src, dst, symlinks = False, ignore = None):
  if not os.path.exists(dst):
    os.makedirs(dst)
    shutil.copystat(src, dst)
  lst = os.listdir(src)
  if ignore:
    excl = ignore(src, lst)
    lst = [x for x in lst if x not in excl]
  for item in lst:
    s = os.path.join(src, item)
    d = os.path.join(dst, item)
    if symlinks and os.path.islink(s):
      if os.path.lexists(d):
        os.remove(d)
      os.symlink(os.readlink(s), d)
      try:
        st = os.lstat(s)
        mode = stat.S_IMODE(st.st_mode)
        os.lchmod(d, mode)
      except:
        pass # lchmod not available
    elif os.path.isdir(s):
      copytree(s, d, symlinks, ignore)
    else:
      shutil.copy2(s, d)



B=0

def openBrowser(isproxy,ip,port,ppassword,pusername):
    global B
    proxy = {'host': ip, 'port': port, 'usr': pusername, 'pwd': ppassword}
    fp = webdriver.FirefoxProfile()
    if isproxy:
        fp.set_preference("extensions.autoauth.currentVersion", "3.1.1")
        fp.set_preference('network.proxy.type', 1)
        fp.set_preference('network.proxy.http', proxy['host'])
        fp.set_preference('network.proxy.http_port', int(proxy['port']))
        fp.set_preference('network.proxy.ssl', proxy['host'])
        fp.set_preference('network.proxy.ssl_port', int(proxy['port']))
        fp.set_preference('network.proxy.no_proxies_on', 'localhost, 127.0.0.1')
        credentials = '{usr}:{pwd}'.format(**proxy)
        credentials = b64encode(credentials.encode('ascii')).decode('utf-8')
        print credentials
        fp.set_preference('extensions.closeproxyauth.authtoken', credentials)
        fp.set_preference("network.proxy.socks_username", proxy['usr'])
        fp.set_preference("network.proxy.socks_password", proxy['pwd'])
    fp.set_preference('media.peerconnection.enabled', False)  # I also tried True, 1 - with and without quotes
    #fp.set_preference('geo.enabled', False)
    fp.set_preference('media.gmp-widevinecdm.enabled', False)
    fp.set_preference('media.gmp-gmpopenh264.enabled', False)
    
    fp.update_preferences()
    B = webdriver.Firefox(firefox_profile=fp)
    B.install_addon(os.path.join(os.getcwd(),"autoauth@efinke.com.xpi"),True)
    try:
        B.get("about:addons")
    except:
        pass
    B.get("https://mail.google.com")
    time.sleep(0.3)
    try:
        B.find_element_by_id("username").send_keys(proxy['usr'])
        time.sleep(0.1)
        B.find_element_by_id("pw").send_keys(proxy['pwd'])
        time.sleep(0.1)
        B.find_element_by_id("submit").click()    
        time.sleep(2)
    except:
        print "No need to,ogin to Proxy server"
    try:
        B.find_element_by_id('warningButton').click()
    except:
        pass 
    copytree(fp.path,"Profiles\\"+proxy['usr'] , symlinks=False, ignore=None)
    return B
    


headers,data = read_xlsx("input.xlsx")


startIndex = 0
if DEBUG:
    startIndex=11
for row,d in enumerate(data[startIndex:]):
    print "--------------------------------------------------------------------"
    print "Running ",row," row"
    login = d['Email']
    password = d['Gmail Password']
    reqemail = d['Gmail Recovery']
    city = d['City'].split(",")
    if type(d['Edit Amount']) == float:
        d['Edit Amount'] = str(int(d['Edit Amount']))
    repeat_times = d['Edit Amount'].split("-")
    if len(repeat_times)>1:
        repeat_times = random.randint((int(repeat_times[0])),int(repeat_times[-1])+1)
    else:
        repeat_times = int(repeat_times[-1])
    
    ip = d['IP']
    port = int(d['Port'])
    pusername = d['Username']
    ppassword = d['Password']
    randomwait = d['Time Delay'].split("-")
    
    print login,password
    print "Edit amount for each gmail account is:", repeat_times
    print ip,port
    print pusername,ppassword
    print randomwait
    
    
    
    print ("Opening Browser")
    if DEBUG:
        B = openBrowser(False,ip,port,ppassword,pusername)
    else:
        B = openBrowser(True,ip,port,ppassword,pusername)
    
    print ("Login to Gmail")
    loginGmail(B,login,password,reqemail)
    
    
    for ci in city:
        print "\nCurrent city is:",ci
        print ("Open Google map")
        x = openGoogleMap(B,ci,repeat_times)
        #break
        print "Aftre checking : EDIT AMOUNT IS:",x
        if x is None:
            print "No new places to review right now"
        else:
            print ("Answer on questions")
            try:
                repeat_times = questionAnswers(B,repeat_times,randomwait)
                if repeat_times==0:
                    print "Done Edit amount!"
                    break
            except Exception as e:
                print e
        print ("Next City")
    print ("Closing Browser")
    
    #break
    B.quit()

print "FINISH!"    