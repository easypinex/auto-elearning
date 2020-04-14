# coding=utf-8
# ----------------
from selenium import webdriver
from time import sleep
from lxml import etree
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities 
import requests
from random import randint
import json
import traceback
import logging
import configparser
from datetime import datetime
import re
from threading import Thread

from PIL import Image#
import numpy as np#
import os#
from subprocess import check_output,STDOUT#
import codecs#


class autoElearning():
    driver = None
    detection = False
    logging = None


    def run(self):
        ocr = True
        debug = False
        config = configparser.ConfigParser()
        config.read('setting.properties',encoding='big5')
        BigExam = str(config['default']['BigExam']) == 'True'
        BigExamUrl = str(config['default']['BigExamUrl'])
        BigExamAnsListUrl = str(config['default']['BigExamAnsListUrl'])
        learnTrueTime = str(config['default']['learnTrueTime']) == 'True'
        learnRatio = float(config['default']['learnRatio'])
        ansOneMinTime = int(config['default']['ansOneMinTime'])
        ansOneMaxTime = int(config['default']['ansOneMaxTime'])
        autoExam = str(config['default']['autoExam']) == 'True'
        
        self.loadLog()
        web = u'員工' if not ocr else u'易學網'
        ac = input(u"請輸入{}網站帳號:".format(web))
        pw = input(u"請輸入{}網站密碼:".format(web))
        # ac = 'hb15358'
        # pw = '!qazxsw2'
        for i in range(30):
            print("")
            
        chrome_options = webdriver.ChromeOptions()
        #Disable Audio
        chrome_options.add_argument("--mute-audio")
        #Flash Player Allow
        prefs = {"profile.default_content_setting_values.plugins": 1,
                "profile.content_settings.plugin_whitelist.adobe-flash-player": 1,
                "profile.content_settings.exceptions.plugins.*,*.per_resource.adobe-flash-player": 1,
                "PluginsAllowedForUrls": "http://elearning.hncb.com.tw:82"}
        chrome_options.add_experimental_option("prefs",prefs)
        #Console Log
        consoleLoader = DesiredCapabilities.CHROME
        consoleLoader['loggingPrefs'] = { 'browser':'ALL'}
        try:
            chromePath = os.path.abspath(open('chromePath.txt').readlines()[0])
            chrome_options.binary_location =chromePath
            driver = webdriver.Chrome('chromedriver.exe',chrome_options=chrome_options,desired_capabilities=consoleLoader)
        except:
            input(u'請至 https://sites.google.com/a/chromium.org/chromedriver/downloads 找尋符合的 ChromeDriver 並取代 (Enter)')
            input(u'請確認chromePath.txt路徑正確 (Enter)')
            chrome_options.binary_location = open('chromePath.txt').readlines()[0]
            driver = webdriver.Chrome('chromedriver.exe',chrome_options=chrome_options,desired_capabilities=consoleLoader)
              
        driver,qs = self.logging(driver,ac,pw,hide=True,ocr = ocr)
        
        
        
        if BigExam:
            answerList = self.getHaveAnswer(qs,driver,BigExamAnsListUrl,BigExam=True)
            open('ans.txt','w',encoding='utf-8').write(str(answerList))
            if(not autoExam):
                return 
            noSample = True
            while noSample:
                driver.get(BigExamUrl)
                driver.find_element_by_id("startBtn").click()
                try:
                    driver.switch_to_alert().accept()
                except:
                    pass
                examHtml = driver.page_source
                lastLen = len(answerList)
                answerList,ansList,noSample,startNum = self.getAns(qs,examHtml,answerList,BigExam)
                #強制等待
                sleep(2)
                driver.implicitly_wait(10)
                allHavaAns,score = self.clickAnswerSubmit(driver,ansList,startNum,ansOneMinTime,ansOneMaxTime)
                beforeLen = lastLen
                if '100' in score or allHavaAns:
                    noSample = False
                if noSample:
                    answerList = self.getHaveAnswer(qs,driver,BigExamAnsListUrl,BigExam=True)
            return
        
        
        
        
        classInfo = self.getClassInfo(driver.page_source)
        print("")
        for i,myclass in enumerate(classInfo):
            print(i,myclass['name'])
        print("")
        logging.debug("classInfo = "+str(classInfo))
        autoElearning.logging = logging
        userChoose = input(u"請輸入需要上課課程 並以,分隔，全自動請輸入 all\n").split(',')
        debug = '-f' in userChoose
        
        if userChoose[0] in ['all','ALL','All']:
            userChoose = [str(i) for i in range(len(classInfo))]
        logging.debug("userChoose = "+str(userChoose))
        for i,myclass in enumerate(classInfo):
            if (not myclass['ispass'] or myclass['score'] < 100) and str(i) in userChoose:
                while True:
                    try:
                        logging.info(u"進入課程:{} ({}/{})".format(myclass['name'],i+1,len(classInfo)))
                        now = -1
                        doSomething = True
                        while doSomething:
                            doSomething = False
                            logging.info(u'讀取子課程...')
                            classBranchInfo = self.getClassBranchInfo(driver,myclass)
                            logging.debug("classBranchInfo = "+json.dumps(classBranchInfo,indent=2))
                            goToClasses,skipList = self.getNeedTolClass(classBranchInfo)
                            now+=1
                            while now in skipList:
                                now+=1
                            for goToClass in goToClasses:
                                logging.debug("goToClass:"+str(goToClass['name']))
                                logging.debug("goToClass="+json.dumps(goToClass,indent=2))
                                if goToClass['type'] == 'scorm': ##上課
                                    doSomething = True
                                    logging.info(u'\t進入子課程:{} ({}/{}) ...'.format(goToClass['name'],now+1,len(classBranchInfo)))
                                    self.learner(driver,goToClass,learnTrueTime,learnRatio)
                                elif goToClass['type'] == 'exam' and autoExam: ##考試
                                    doSomething = True
                                    logging.info(u'\t考試:{} ({}/{})...'.format(goToClass['name'],now+1,len(classBranchInfo)))
                                    noSample = True
                                    noAnsLimit = 10
                                    answerList = self.getHaveAnswer(qs,driver,goToClass['url'])
                                    beforeLen = -1
                                    while noSample:
                                        url = 'https://elearning.hncb.com.tw/WcmsModules/OnLineClass/'+goToClass['url']
                                        driver.get(url)
                                        driver.find_element_by_id("startBtn").click()
                                        try:
                                            driver.switch_to_alert().accept()
                                        except:
                                            pass
                                        examHtml = driver.page_source
                                        lastLen = len(answerList)
                                        answerList,ansList,noSample,startNum = self.getAns(qs,examHtml,answerList)
                                        #強制等待
                                        sleep(2)
                                        driver.implicitly_wait(10)
                                        allHavaAns,score = self.clickAnswerSubmit(driver,ansList,startNum,ansOneMinTime,ansOneMaxTime)
                                        beforeLen = lastLen
                                        if '100' in score or allHavaAns:
                                            noSample = False
                                        if noSample:
                                            answerList = self.getHaveAnswer(qs,driver,goToClass['url'])
                                elif goToClass['type'] == 'questionnaire' and autoExam:##問卷
                                    doSomething = True
                                    logging.info(u'\t填寫問卷中 ...')
                                    try:
                                        self.questionnaire(driver,qs,goToClass)
                                    except:
                                        logging.info(u'\t填寫問卷失敗 ! (可能已填過)')
                        break
                    except:
                        logging.error(traceback.format_exc())
                        if not debug:
                            logging.error('error .. skip')
                            break
                    
                        
            else:
                logging.info(u"跳過可結案課程或非選擇:{} ({}/{})".format(myclass['name'],i+1,len(classInfo)))
        logging.info(u"完成所有指定課程! Enter 結束!")
        input("")
    
    def getClassInfo(self,html):
        tree = etree.HTML(html)
        table = tree.xpath('//div[@id="ctl00_ContentPlaceHolder1_PageLayout1_ctl02_panel1"]')[0]
        myclasses = table.xpath('//fieldset/table[@class="table"]/tbody')[0][1:]
        classInfo = []
        
        for myclass in myclasses:
            name = myclass.xpath('td/a')[0].text
            href = myclass.xpath('td/a')[0].get('href')
            try:
                score = int(myclass.xpath('td/div/span/table/tbody/tr[3]/td[2]/text()')[0].split('目前平均成績:')[-1].split(')')[0])
            except:
                score = 0
            coursePK = href[href.find('coursePK='):]
            coursePK = coursePK[:coursePK.find('&')]
            classPK = href[href.find('classPK='):]
            classPK = classPK[:classPK.find('&')]
            ispass = 'green.gif' in myclass.xpath('td/img')[1].get('src')
            classInfo.append({'name':name,
                              'coursePK':coursePK,
                              'classPK':classPK,
                              'ispass':ispass,
                              'score':score})
        return classInfo
    
    
    def getClassBranchInfo(self,driver,myclass):
        coursePK = myclass['coursePK'].replace('coursePK','CoursePK')
        classPK = myclass['classPK'].replace('classPK','ClassPK')
        classUrl = 'https://elearning.hncb.com.tw/WcmsModules/OnLineClass/TopMain.aspx?{}&{}&Mode=S'.format(coursePK,classPK)
        driver.get(classUrl)
        try:
            WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "iframeTips")))
            driver.switch_to_frame("iframeTips")
            driver.find_element_by_id("btnCancelDialog").click()
            driver.switch_to.default_content()
        except:
            pass
        driver.switch_to.frame(driver.find_element_by_id("iframe1"))
        html = driver.page_source
        
        tree = etree.HTML(html)
        forceClass = False
        classes = tree.xpath('//div[@id="ResourceList1_StudentArea"]/table/tbody')[0][1:]
        ClassBranchInfo = []
        for myclass in classes[-1::-1]:
            type = myclass.xpath('td/div')[0].get('class')
            name = myclass.xpath('td/a')[0].text
            onclick = myclass.xpath('td/a')[0].get('onclick')
            finished = myclass.xpath('td/table/tbody/tr/td[2]/text()')[0].strip()
            infoBox = tree.xpath('//div[@id="ClassEndCondition1_divShow"]/text()')[0].strip()
            infoBox = infoBox.split('分鐘')
            needMin = int(infoBox[0].replace('閱讀完所有教材，且總時數達',''))
            nowMin = int(infoBox[1].replace('以上,目前進度為',''))
            if '100' == finished.split('%')[0].replace('完成度 : ',''):
                finished = True
            elif '完成度' not in finished and '未完成' not in finished and '完成' in finished and type != 'scorm':
                finished = True
            elif '100.00' in finished:
                finished = True
            else:
                finished = False
            url = ''
            if onclick != '': 
                if type == 'scorm':
                    onclick1 = onclick.replace("GoToClassRoom('","").replace("')","")
                    onclick1 = onclick1.split(',')
                    url = onclick1[0].replace("'","")
                    cacheKey = onclick1[1].replace("'","")
                    classID = onclick1[2].replace("'","")
                    rid = onclick1[3].replace("'","")
                    catch = url.replace('http://elearning.hncb.com.tw:82/','')
                    catch = 'Class_' + catch.split('/')[0]
                    # https://elearning.hncb.com.tw/WcmsModules/OnLineClass/OpenScormReader.aspx?classID=B202003050001&RID=rc-2cfb49f3-e827-4d1d-a8e3-74a83a96473a&cacheKey=Class_c82ae038-a396-41d8-9fd2-e754b4a4fb74&readerURL=http%3A//elearning.hncb.com.tw%3A82/69a25b72-0777-4d0d-aee9-2d651db5989b/rcdata/privatecourse/rc-2cfb49f3-e827-4d1d-a8e3-74a83a96473a/%24startup
                    url = 'https://elearning.hncb.com.tw/WcmsModules/OnLineClass/OpenScormReader.aspx?classID={classID}&RID={rid}&cacheKey={cacheKey}&readerURL={url}'.format(
                        classID = classID,
                        rid = rid,
                        cacheKey = cacheKey,
                        url = url)
                elif type == 'exam':
                    url = onclick[onclick.find('ExamModule'):]
                    url = url[:url.find("'")]
                    if len(url) > 0:
                        forceClass = nowMin<needMin
                        if forceClass:
                            logging.info(u'\t分鐘數不足 (已讀 {} 分鐘 , 需要 {} 分鐘) ,且已有考試,強制啟用閱讀'.format(nowMin,needMin))
                elif type == 'questionnaire':
                    url = onclick[onclick.find('SCROMWrapper'):]
                    url = url[:url.find("'")]
            ClassBranchInfo.append({'type':type,
                                    'name':name,
                                    'url':url,
                                    'onclick':onclick,
                                    'force':forceClass,
                                    'finished':finished,
                                    'nowMin':nowMin,
                                    'needMin':needMin})
        return ClassBranchInfo
    
    def getNeedTolClass(self,classBranchInfo):
        skipNum = 0
        skipList = []
        lastClass = []
        beenForce = False
        for myclass in classBranchInfo[-1::-1]:
            if myclass['url'] != '':
                if myclass['force'] and myclass['type'] == 'scorm':
                    beenForce = True
                    lastClass.append(myclass)
                else:
                    if beenForce or not myclass['finished']:
                        lastClass.append(myclass)
                    else:
                        skipList.append(skipNum)
                        skipNum+=1
                        
        return lastClass,skipList
    
    def learner(self,driver,classBranch,TrueLearn=False,learnRatio=1.5):
        FlashPlayer = False
        
        logging.debug("into:"+classBranch['url'])
        driver.get_log('browser')
        loaded = False
        while not loaded:
            try:
                driver.get(classBranch['url'])
                logging.debug("wait by contentframe ...")
                WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, "contentframe")))
                loaded = True
            except:
                logging.debug("page is error? try again...")
        logging.debug("wait by contentframe ...")
        WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.ID, "menuframe")))
        learnAll = classBranch['nowMin'] < classBranch['needMin']
        if learnAll:
            script = ''.join(open('command.txt').readlines())
        else:
            script = ''.join(open('command-com.txt').readlines())
        sleep(3)
        logging.debug("switch_to_frame menuframe...")
        driver.switch_to_frame(driver.find_element_by_id("menuframe"))
        WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.XPATH, '//ul[@class="k-group k-treeview-lines"]')))
        html = driver.page_source
        tree = etree.HTML(html)
        classes = len(tree.xpath('//ul[@class="k-group k-treeview-lines"]')[0])
        min = (classes*3) // 60
        sec = (classes*3) - 60*min
        if not TrueLearn:
            logging.info(u"\t\t上課中...視訊量:{} 所需約:{} 分 {} 秒 ... ".format(classes,min,sec))
        else:
            logging.info(u"\t\t上課中...視訊量:{}".format(classes))
        self.waitConsole(driver,'APIAdapter initialized')
#         driver.execute_script('$(window).unbind("blur focusout");')
        learningTime,learnMin,learnSec,TotalSec = self.getLearnTime(driver,learnRatio)
        thisScript = script.format(min = str(learnMin).zfill(2),sec=str(learnSec).zfill(2),rnd=str(randint(0,99)).zfill(2)) if learnAll else script
        logging.debug('thisScript={}'.format(thisScript))
        if TrueLearn:
            logging.debug('sleep({})'.format(TotalSec))
            autoElearning.driver = driver
            autoElearning.detection = True
            sleep(TotalSec)
            autoElearning.detection = False
            sleep(0.5)
            
        logging.debug("exec script...:"+thisScript)
        driver.execute_script(thisScript)
        if learnAll:
            es = driver.find_elements_by_xpath('//div/span[@class="k-in"]')
        else:
            es = driver.find_elements_by_xpath('//span[@class="k-in"]/img[@src="/images/not attempted.gif"]')
            es += driver.find_elements_by_xpath('//span[@class="k-in"]/img[@src="/images/incomplete.gif"]')
        for e in es:
            driver.get_log('browser')
            e.click()
            self.waitConsole(driver,'APIAdapter initialized')
            
            learningTime,learnMin,learnSec,TotalSec = self.getLearnTime(driver,learnRatio)
            thisScript = script.format(min = str(learnMin).zfill(2),sec=str(learnSec).zfill(2),rnd=str(randint(0,99)).zfill(2)) if learnAll else script
            logging.debug('thisScript={}'.format(thisScript))
            if TrueLearn:
                logging.debug('sleep({})'.format(TotalSec))
                autoElearning.driver = driver
                autoElearning.detection = True
                sleep(TotalSec)
                autoElearning.detection = False
                sleep(0.5)
            logging.debug("exec script...:"+thisScript)
            driver.get_log('browser')
            driver.execute_script(thisScript)
            self.waitConsole(driver,'API.LMSSetValue(cmi.core.lesson_status, completed)')
            sleep(1)
            
        sleep(2)
        logging.debug("switch_to.default_content ...")
        driver.switch_to.default_content()
        logging.debug("switch systemframe ...")
        driver.switch_to_frame(driver.find_element_by_id("systemframe"))
        logging.debug(u"click 儲存並離開 ...")
        driver.find_element_by_xpath("//*[@value='儲存並離開']").click()
        logging.debug(u"wait... 關閉 ...")
        btn = WebDriverWait(driver,180).until(EC.element_to_be_clickable((By.CLASS_NAME, "button")))
        logging.debug(u"click... 關閉 ...")
        btn.click()
        logging.debug("learn exit!")

    def waitConsole(self,driver,waitKey,timeout=15):
        timeCount = 0
        waited = False
        while timeCount < timeout:
            for entry in driver.get_log('browser'):
                if waitKey in entry['message']:
                    logging.debug("find:{}".format(entry['message']))
                    return entry['message']
            sleep(1)
            timeCount+=1
        return None
            
    def getLearnTime(self,driver,learnRatio):
        driver.get_log('browser')
        learningTime = '00:00'
        while learningTime == '00:00' or learningTime == None or '' == learningTime:
            try:
                driver.execute_script("""console.log("auto-learn:"+parent.window.$('#contentframe').contents().find('#myElement_controlbar_duration').text());""")
                learningTime = self.waitConsole(driver, 'auto-learn:')
                driver.get_log('browser')
                if learningTime != None:
                    learningTime = learningTime.split('"')[1].replace('auto-learn:','')
                if learningTime == '00:00' or learningTime == None '' == learningTime:
                    sleep(0.5)
            except:
                logging.error(traceback.format_exc())
        logging.debug('learningTime={}'.format(learningTime))
        learnMin = int(learningTime.split(":")[0])
        learnSec = int(learningTime.split(":")[1]) + learnMin*60
        learnSec = int(float(learnSec)*learnRatio)
        TotalSec = learnSec
        learnMin = learnSec//60
        learnSec = learnSec-learnMin*60
        logging.debug('learnMin={},learnSec={},TotalSec={}'.format(learnMin,learnSec,TotalSec))
        
        return learningTime,learnMin,learnSec,TotalSec
        
    def getHaveAnswer(self,qs,driver,url,BigExam = False):
        logging.info(u"\t採樣中...")
        examPK = url[url.find('ExamPK='):]
        examPK = examPK[:examPK.find('&')]
        answerList = []
        answerListFinal = []
        if not BigExam:
            examListUrl = "https://elearning.hncb.com.tw/WcmsModules/OnLineClass/ExamModule/QueryScore.aspx?" + examPK
            text = qs.get(examListUrl).text
            answerUrlList = self.getAnsUrlList(text)
        else:
            examListUrl = url
            text = qs.get(examListUrl).text
            answerUrlList = self.getAnsUrlList_BigExam(text)
            
        for ansurl in answerUrlList:
            if not BigExam:
                text = requests.get(ansurl).text
            else:
                driver.get(ansurl)
                text = driver.page_source
            answerList = answerList + self.examParser(text,BigExam)
            for ans in answerList:
                if ans not in answerListFinal:
                    answerListFinal.append(ans)
#             answerList = list(set(answerList))
#         for DEBUG
#         for ans,text in answerListFinal:
#             print ans,text
        logging.info(u"\t收集答案數:"+str(len(answerListFinal)))
        return answerListFinal
            
    def getAnsUrlList(self, html):
        anslist = []
        tree = etree.HTML(html)
        table = tree.xpath('//table[@width="90%"]')[1][1:]
        for e in table:
            td = e.xpath('td')[-1]
            lastUrl = etree.tostring(td).decode('utf-8').split('window.open(\'')[-1].split('\');" ')[0]
            lastUrl = lastUrl.replace('&amp;', "&")
            anslist.append('https://elearning.hncb.com.tw/WcmsModules/OnLineClass/ExamModule/' + lastUrl)
        return anslist
    
    def getAnsUrlList_BigExam(self, html):
        anslist = []
        tree = etree.HTML(html)
        table = tree.xpath('//table[@width="100%" and @cellpadding="1"]')[0][1:]
        for e in table:
            td = e.xpath('td')[-1]
            lastUrl = etree.tostring(td).decode('utf-8').split('window.open(\'')[-1].split('\',\'\',\'')[0]
            lastUrl = lastUrl.replace('&amp;', "&")
            anslist.append('https://elearning.hncb.com.tw/_service/omega/pretest/' + lastUrl)
        return anslist
    
    def examParser(self, html,BigExam=False):
        mlist = []
        tree = etree.HTML(html)
        if not BigExam:
            table = tree.xpath('//*[@id="form1"]/div/table[@width="100%"][1]/tr[2]/td/table[3]/tr/td')
            
            for box in table:
                q = box.xpath('span[1]')[0].xpath('normalize-space()')
                if u'是非題,共' in q or u'單選題,共' in q or u'複選題,共' in q:
                    continue
                ansbox = box.xpath('table/tr')
                ans = ""
                if len(ansbox) <= 3:
                    # 是非題
                    for i in range(len(ansbox[1:])):
                        if len(ansbox[1:][i].xpath('td[2]/img')) > 0:
                            ans = "2"
                        else:
                            ans = "1"
                    firstOptionText = self.cleanText(''.join(ansbox[1:][0].xpath('td[3]/text()')))
    #                 list.append(ans+"\t\t"+q)
                else:
                    # 選擇題 or 複選題
                    for i in range(len(ansbox[1:])):
                        if len(ansbox[1:][i].xpath('td[2]/img')) > 0:
                            if type(ans) is str:
                                if ans == "": #沒有答案
                                    ans = str(i + 1)  # +','+ansbox[1:][i].xpath('td[3]')[0].xpath('normalize-space()')
                                    firstOptionText = self.cleanText(''.join(ansbox[1:][0].xpath('td[3]/text()')).strip())
                                else: # 已經有答案  代表複選題
                                    ans = [ans,str(i + 1)]
                            elif type(ans) is list: #複選題
                                ans.append(str(i + 1))
    #             print firstOptionText.encode('big5','ignore')
                mlist.append((ans, self.cleanText(q.strip()),firstOptionText))
        else:
            table = tree.xpath('/html/body/table/tbody/tr[2]/td/table[@class="table" and @width="100%" and not(@cellpadding="3")]/tbody/tr')
            for box in table:
                if len(box.xpath('td[@class="thead"]')) > 0:
                    continue
                q = box.xpath('td[@class="tdrowbody"]/table[@class="tdrowbody"]/tbody/tr/td[@class="topicCaption"]')[0].xpath('normalize-space()')
                q = ''.join(q.split('.')[1:]).split(u' (題目配分：')[0]
                ansbox = box.xpath('td[@class="tdrowbody"]//tr[@bgcolor="#EEEEEE"]')
                ansText = ""
                if len(ansbox) <= 2:
                    # 是非題
                    for i in range(len(ansbox)):
                        if len(ansbox[i].xpath('td[2]/img')) > 0:
                            ansText = ''.join(ansbox[i].xpath('td[3]/text()'))
                else:
                    # 選擇題 or 複選題
                    for i in range(len(ansbox)):
                        if len(ansbox[i].xpath('td[2]/img')) > 0:
                            if type(ansText) is str:
                                if ansText == "": #沒有答案
                                    ansText = ''.join(ansbox[i].xpath('td[3]/text()'))
                                else: # 已經有答案  代表複選題
                                    ansText = [ansText,''.join(ansbox[i].xpath('td[3]/text()'))]
                            elif type(ansText) is list: #複選題
                                ans.append(''.join(ansbox[i].xpath('td[3]/text()')))
                mlist.append((ansText, self.cleanText(q.strip())))
        return mlist
    
    
    def getAns(self,qs,examHtml,answerList,BigExam=False):
        tree = etree.HTML(examHtml)
        table = tree.xpath('//div[@id="OptContentDIV"]/table/tbody/tr/td/div')
        no = 0
        ansList = []
        noSample = False
        
        if not BigExam:
            for i,examqution in enumerate(table):
                no += 1
                examqutionText = (''.join(examqution.xpath('text()')).strip()).split('分)')[-1].strip()
                examFirstOptionText = ''.join(examqution.xpath('div/table/tbody/tr[1]/td/label/text()')).strip()
                id = re.sub('.*(\d+).*',r'\1',examqution.xpath('span')[0].get('id')[0:-2]).zfill(2)
                if i == 0:
                    startNum = id
                find = False
                for ans, q, firstOptionText in answerList:
                    if self.cleanText(examqutionText) in q and firstOptionText in self.cleanText(examFirstOptionText):
                        ansList.append(ans)
    #                     print examFirstOptionText.encode('big5','ignore')
    #                     print firstOptionText.encode('big5','ignore')
                        find = True
                        break
                if not find:
                    ansList.append(None)
                    noSample = True
        else:
            for i,examqution in enumerate(table):
                no += 1
                find = False
                examqutionText = (''.join(examqution.xpath('text()')).strip()).split('points) ')[-1].strip()
                ##div/table/tbody/tr[1]/td/label/text()
                for j,option in enumerate(examqution.xpath('div/table/tbody/tr/td/label')):
                    examOptionText = ''.join(option.xpath('text()'))
                    id = re.sub('.*(\d+).*',r'\1',examqution.xpath('span')[0].get('id')[0:-2]).zfill(2)
                    if i == 0:
                        startNum = id
                    for ansText, q in answerList:
                        if self.cleanText(examqutionText) == self.cleanText(q):
                            if self.cleanText(examOptionText) == self.cleanText(ansText):
                                ansList.append(str(j+1))
                                find = True
                                break
                if not find:
                    ansList.append(None)
                    noSample = True
        return answerList,ansList,noSample,startNum
    
    def clickAnswerSubmit(self,driver,answer,startNum,minSleepForOne,maxSleepForOne,hide=True):
        logging.info(u'\t輸入答案: {} ...'.format(str(answer)))
        try:
            driver.switch_to_alert().accept()
        except:
            pass
        driver.execute_script('$(window).unbind("blur focusout");')
        allHavaAns = True
        ansList = []
        logging.info(u'\t答題數:{}'.format(len(answer)))
        for i,ans in enumerate(answer):
            if ans == None:
                ans = str(randint(1,2))
                allHavaAns = False
            ansList.append(ans)
            sleep(randint(minSleepForOne,maxSleepForOne))
            if type(ans) is str:
                fullId = 'QuestionItemList_ctl{}_isAnswer_{}'.format(str(int(startNum)+i).zfill(2),int(ans)-1)
#                 el = driver.find_elements_by_id(fullId)[-1] 
                driver.find_element_by_xpath("//input[@id='{}']".format(fullId)).click()
#                 el = driver.find_element_by_xpath("//input[@id='{}']".format(fullId))
#                 action = webdriver.common.action_chains.ActionChains(driver)
#                 action.move_to_element(el).perform()
#                 action.move_to_element_with_offset(el,6,6)
#                 action.click().perform()
            elif type(ans) is list:
                for a in ans:
                    fullId = 'QuestionItemList_ctl{}_isAnswer_{}'.format(str(int(startNum)+i).zfill(2),int(a)-1)
                    driver.find_element_by_xpath("//input[@id='{}']".format(fullId)).click()
#                     driver.find_element_by_xpath("//input[@id='{}' and @type='radio']".format(fullId)).click()
#                     el = driver.find_element_by_id('QuestionItemList_ctl{}_isAnswer_{}'.format(str(int(startNum)+i).zfill(2),int(a)-1))
#                     action = webdriver.common.action_chains.ActionChains(driver)
#                     action.move_to_element_with_offset(el,6,6)
#                     action.click()
#                     action.perform()
        try:
            driver.maximize_window()
        except:
            logging.debug('fail to max ,skip:'+traceback.format_exc())
        driver.find_element_by_name("saveBtn").click()
        if hide :
            try:
                driver.set_window_position(-10000,0)
            except:
                logging.debug('fail to hide , skip:'+traceback.format_exc())
#         saveBtn = driver.find_element_by_name("saveBtn")
#         action = webdriver.common.action_chains.ActionChains(driver)
#         action.move_to_element_with_offset(saveBtn,6,6)
#         action.click()
#         action.perform()
        try:
            driver.switch_to_alert().accept()
#             if(allHavaAns):
#                 driver.switch_to_alert().dismiss()
#                 ansListText = ""
#                 for j,ansNum in enumerate(ansList):
#                     ansListText+= u'題:{} 答: {}\t'.format(j+1,str(ansNum))
#                     if j+1 % 3 == 0:
#                         ansListText+='\n'
#                 print ansListText
#                 input(u'有答案無法填寫，請手動點擊答案！ 按下Eneter放大瀏覽器\n')
#                 try:
#                     driver.maximize_window()
#                 except:
#                     logging.debug('fail to max ,skip:'+traceback.format_exc())
#                 input(u'有答案無法填寫，請手動點擊答案並繳卷繼續！\n')
#             else:
#                 driver.switch_to_alert().accept()
        except:
            pass
        logging.info(u'\t繳卷成功!')
        score = driver.find_element_by_xpath('//span[@id="Score"]').text.split('分')[0]
        logging.info(u'\t分數:{}'.format(score))
        
        return allHavaAns,score
    
    def questionnaire(self,driver,qs,myclass):
        PK = myclass['url'][myclass['url'].find('PK='):]
        PK = PK[:PK.find('&')]
        PK = PK.replace('PK=','')
        self.goToHome(driver)
        html = driver.page_source
        userPK = html[html.find('UserPK='):]
        userPK = userPK[:userPK.find('&')]
        userPK = userPK.replace('UserPK=',"")
        driver.get('https://elearning.hncb.com.tw/_service/omega/survey2/DoClassSurveyforMobile.asp?PK={}'.format(PK))
        for i in range(4):
            driver.switch_to_alert().accept()
        html = driver.page_source
        tree = etree.HTML(html)
        radios = tree.xpath('//*[@type="radio"]')[0:-1:5]
        data = {'surveyPK': PK,
                'for': 'vclass',
                'userPK': userPK}
        for i,radio in enumerate(radios):
            data[radio.get(b'name')] = radio.get(b'value')
            data["qid"+radio.get(b'name').replace("Q","")] = str(i+1)
        rs = qs.post('https://elearning.hncb.com.tw/_service/omega/survey2/ActionSurveyScoreMobile.asp',data = data)
        
    
        
    
                
    
    
    
    
    
    def logging(self,driver,ac,pw,hide = True,ocr = False):
        while True:
            try:
                web = u'員工' if not ocr else u'易學網'
                logging.info(u'登入中...')
                if not ocr:
                    qs = requests.session()
                    #登入
                    data = {'LoginDomain':'huanan.com.tw',
                                'LoginId':ac,
                                'LoginPassword':pw}
                    logging.info(u'登入員工網站...')
                    qs.post("http://staffnew.hncb.com.tw/Account/ADLogin", data)
                    #易學網
                    logging.info(u'進入易學網...')
                    qs.post("http://staffnew.hncb.com.tw/Home/OtherSystem/d3cbdf9d-966e-4009-a2c3-93ff29b90fd6")
                    
                    driver.get('https://elearning.hncb.com.tw/WcmsModules/Portal/FullFrameLogin/index.aspx')
                    driver.delete_all_cookies()
                    for cookie in qs.cookies:
                        if ("elearning" in cookie.domain):
                            newCookie = {'domain':cookie.domain,
                                               'name':cookie.name,
                                               'value':cookie.value,
                                               'path':"/",
                                               'expires':None}
                            driver.add_cookie(newCookie)
                else:
                    driver.get('https://elearning.hncb.com.tw/WcmsModules/Portal/FullFrameLogin/index.aspx')
                    sleep(0.5)
                    img = driver.find_elements_by_tag_name('img')[0]
                    fileName = 'screenshot.png'
                    self.get_captcha(driver, img, fileName)
                    im = Image.open(fileName)
                    im = np.array(im)
                    im = self.cleanCodeImg(im)
                    im = Image.fromarray(im)
                    im.save(fileName)
                    code = self.image_to_string('screenshot.png', True, '')
                    driver.find_elements_by_id("UserID")[0].send_keys(ac)
                    driver.find_elements_by_id("Password")[0].send_keys(pw)
                    driver.find_elements_by_id("captcha")[0].send_keys(code)
                    
                if hide :driver.set_window_position(-10000,0)
                if not ocr:
                    driver.get('https://elearning.hncb.com.tw/Index.aspx')
                if ocr:
                    driver.find_elements_by_id("button")[0].click()
                    os.remove(fileName)
                qs = requests.session()
                for cookie in driver.get_cookies():
                    qs.cookies.set(cookie['name'], cookie['value'])
                driver.switch_to_frame(driver.find_element_by_name("bottomframe"))
                driver.switch_to_frame(driver.find_element_by_name("display"))

                logging.info(u"登入成功!")
                break
            except Exception as e:
                logging.error(e)
                logging.error(traceback.format_exc())
                a = input(u'登入失敗,是否要重新輸入帳號密碼? (y/n)')
                if a == 'y':
                    ac = input(u"請輸入{}網站帳號:".format(web))
                    pw = input(u"請輸入{}網站密碼:".format(web))
        return driver,qs

    def cleanText(self,text):
        text = text.replace(' ','')
        text = text.replace('\n','')
        return text
    
    
    def goToHome(self,driver):
        driver.get('https://elearning.hncb.com.tw/Index.aspx')
        driver.switch_to_frame(driver.find_element_by_name("bottomframe"))
        driver.switch_to_frame(driver.find_element_by_name("display"))
        return driver
    
    def get_captcha(self,driver, element, path):
        # now that we have the preliminary stuff out of the way time to get that image :D
        location = element.location
        size = element.size
        # saves screenshot of entire page
        driver.save_screenshot(path)
        # uses PIL library to open image in memory
        image = Image.open(path)
    
        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
    
        image = image.crop((left, top, right, bottom))  # defines crop points
        image.save(path)  # saves new cropped image
        
    def cleanCodeImg(self,im):
        try:
            im = np.delete(im,3, axis=2)
        except:
            pass
        for x in range(len(im)):
            for y in range(len(im[x])):
                if im[x][y].sum() < 255*3:
                    im[x][y] = [0,0,0]
        
        return im
    
    def image_to_string(self,img, cleanup=True, plus=''):
        # cleanup为True则识别完成后删除生成的文本文件
        # plus参数为给tesseract的附加高级参数
        command = r'{ocr} {img} {img} {plus} --tessdata-dir "{ocr}\tessdata"'.format(ocr = os.path.join(os.getcwd(),'OCR','tesseract'),img=img,plus = plus)
        check_output(command, shell=True,stderr=STDOUT)  # 生成同名txt文件
        text = ''
        with codecs.open(img + '.txt', "r",encoding='utf-8', errors='ignore') as f:
            text = f.read().strip()
        if cleanup:
            os.remove(img + '.txt')
        text = text.replace(" ", "").replace("'","").replace(",","").replace('-','')
        return text

    def loadLog(self):
        logFormatter = logging.Formatter("[%(asctime)s] %(message)s","%Y-%m-%d %H:%M:%S")
        logFormatter2 = logging.Formatter("[%(levelname)s][%(asctime)s] %(message)s","%Y-%m-%d %H:%M:%S")
        rootLogger = logging.getLogger()
        fileHandler = logging.FileHandler("log.log",'w', encoding='utf-8')
        fileHandler.setFormatter(logFormatter2)
        fileHandler.setLevel(logging.DEBUG)
        rootLogger.addHandler(fileHandler)
        consoleHandler = logging.StreamHandler()
        consoleHandler.setFormatter(logFormatter)
        consoleHandler.setLevel(logging.INFO)
        rootLogger.addHandler(consoleHandler)
        rootLogger.setLevel(logging.DEBUG)
        logging.getLogger("requests").setLevel(logging.WARNING)
        logging.getLogger("selenium").setLevel(logging.WARNING)
class detection_event(Thread):
    def run(self):
        while True:
            if autoElearning.detection:
                driver = autoElearning.driver
                logging = autoElearning.logging
                try:
                    driver.switch_to.default_content()
                    btn = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH, '//input[@onclick="fnTimeIsUpCheck()"]')))
                    if btn.is_displayed():
                        logging.info(u'\t點擊繼續閱讀...')
                        btn.click()
                    driver.switch_to_frame(driver.find_element_by_id("menuframe"))
                except Exception as e:
                    logging.error(e)
                sleep(0.1)
detection_event().start()
autoElearning().run()