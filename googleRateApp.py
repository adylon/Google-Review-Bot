# -*- coding: utf-8 -*-
"""
Created on Tue Oct  9 16:07:18 2018

@author: abrahamy
"""

from selenium import webdriver
import xlrd
import time
from selenium.webdriver.common.keys import Keys
from base64 import b64encode
import shutil
import stat
import os
####################################################################################################
####################################################################################################
####################################################################################################
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
####################################################################################################
####################################################################################################
####################################################################################################
def loginGmail(B,login,password,remail):
    B.get("https://accounts.google.com")
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
####################################################################################################
####################################################################################################
####################################################################################################
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
        print "No need to login to Proxy server"
    try:
        B.find_element_by_id('warningButton').click()
    except:
        pass 
    copytree(fp.path,"Profiles\\"+proxy['usr'] , symlinks=False, ignore=None)
    return B
    
def rateUrlF(B,rateUrl,reviewText):
    B.get(rateUrl)
    time.sleep(4)
    for i in range(5):
        try:
            B.find_elements_by_class_name("section-write-review")[0].click()    
            break
        except:
            time.sleep(1)
    else:
        return None
    time.sleep(2)
    for i in range(5):
        try:
            B.switch_to_frame(B.find_element_by_class_name("goog-reviews-write-widget"))
            break
        except:
            time.sleep(2)
    time.sleep(1)
    for i in range(5):
        try:
            B.find_elements_by_class_name("rating-star")[-1].click()
            break
        except:
            time.sleep(2)
    time.sleep(1)
    B.find_element_by_class_name("review-text").clear()
    B.find_element_by_class_name("review-text").send_keys(reviewText)
    time.sleep(1)
    B.find_elements_by_class_name("quantum-button-inner-box")[-1].click()
    time.sleep(1)
    return True
####################################################################################################
####################################################################################################
####################################################################################################
headers,data = read_xlsx("Review Input.xlsx")
####################################################################################################
####################################################################################################
####################################################################################################
for row,d in enumerate(data[:]):
    print "--------------------------------------------------------------------"
    print "Running ",row," row"
    login = d['Email']
    password = d['Gmail Password']
    reqemail = d['Gmail Recovery']
    
    ip = d['IP']
    port = int(d['Port'])
    pusername = d['Username']
    ppassword = d['Password']
    rateUrl = d['Google My Business URL']
    reviewText = d['Review']
    
    print login,password
    print ip,port
    print pusername,ppassword
    print rateUrl
    print reviewText
    
    
    print ("Opening Browser")
    B = openBrowser(True,ip,port,ppassword,pusername)
    
    print ("Login to Gmail")
    loginGmail(B,login,password,reqemail)
    
    print "Open URL which should be rated"
    j = rateUrlF(B,rateUrl,reviewText)
    if j:
        print "Rated"
    else:
        print "Not Rated!"
    B.quit()
####################################################################################################
####################################################################################################
####################################################################################################
print "--------------------------------------------------------------------"
print "FINISH!"    
####################################################################################################
####################################################################################################
####################################################################################################
