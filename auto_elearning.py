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
from urllib3.exceptions import InsecureRequestWarning
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
from getpass import getpass

# Suppress only the single warning from urllib3 needed.
# requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)
# requests.packages.urllib3.util.ssl_.DEFAULT_CIPHERS += ':AES256-GCM-SHA384'

elearning_domain = 'https://elearning.taiwanlife.com'
class autoElearning():
    driver = None
    detection = False
    loggin = None
    

    def run(self):
        ocr = True
        debug = True
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
        web = u'數位學習平台'
        # ac = input(u"請輸入{}網站帳號:".format(web))
        # pw = getpass(u"請輸入{}網站密碼:".format(web))
        ac = '196693'
        pw = 'eL19941019'

        chrome_options = webdriver.ChromeOptions()
        #Disable Audio
        chrome_options.add_argument("--mute-audio")
        #Disalbe SSL
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_argument('--ignore-certificate-errors')
        #Flash Player Allow
        prefs = {"profile.default_content_setting_values.plugins": 1,
                "profile.content_settings.plugin_whitelist.adobe-flash-player": 1,
                "profile.content_settings.exceptions.plugins.*,*.per_resource.adobe-flash-player": 1,
                "PluginsAllowedForUrls": elearning_domain}
        chrome_options.add_experimental_option("prefs",prefs)
        #Console Log
        consoleLoader = DesiredCapabilities.CHROME
        consoleLoader['loggingPrefs'] = { 'browser':'ALL'}
        try:
            driver = webdriver.Chrome('chromedriver.exe',chrome_options=chrome_options,desired_capabilities=consoleLoader)
        except Exception as e :
            input(u'請至 https://chromedriver.chromium.org/downloads 找尋符合的 ChromeDriver 並取代 (Enter)')
            input(u'請確認chromePath.txt路徑正確 (Enter)')
            chrome_options.binary_location = open('chromePath.txt').readlines()[0]
            driver = webdriver.Chrome('chromedriver.exe',chrome_options=chrome_options,desired_capabilities=consoleLoader)
              
        driver,qs = self.loggin(driver,ac,pw,hide=False)
        
        
        
        if BigExam:
            answerList = self.getHaveAnswer(qs,driver,BigExamAnsListUrl,BigExam=True)
            open('ans.txt','w',encoding='utf-8').write(str(answerList))
            if(not autoExam):
                return 
            noSample = True
            while noSample:
                driver.get(BigExamUrl)
                driver.find_element(By.ID, "startBtn").click()
                try:
                    driver.switch_to_alert().accept()
                except:
                    pass
                examHtml = driver.page_source
                lastLen = len(answerList)
                answerList,ansList,noSample,startNum = self.getAns(driver, qs,examHtml,answerList,BigExam)
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
        
        
        
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.CLASS_NAME, "timeline")))
        classInfo = self.getClassInfo(driver.page_source)
        print("")
        for i,myclass in enumerate(classInfo):
            print(i,myclass['name'])
        print("")
        logging.debug("classInfo = "+str(classInfo))
        autoElearning.loggin = logging
        if not debug:
            userChoose = input(u"請輸入需要上課課程 並以,分隔，全自動請輸入 all\n").split(',')
            debug = '-f' in userChoose
        
        if debug or userChoose[0] in ['all','ALL','All']:
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
                                if goToClass['type'] == '教材': ##上課
                                    doSomething = True
                                    logging.info(u'\t進入子課程:{} ({}/{}) ...'.format(goToClass['name'],now+1,len(classBranchInfo)))
                                    self.learner(driver,goToClass,learnTrueTime,learnRatio)
                                elif goToClass['type'] == '考試' and autoExam: ##考試
                                    doSomething = True
                                    logging.info(u'\t考試:{} ({}/{})...'.format(goToClass['name'],now+1,len(classBranchInfo)))
                                    noSample = True
                                    noAnsLimit = 10
                                    answerList = self.getHaveAnswer(qs,driver,goToClass['onclick'])
                                    beforeLen = -1
                                    
                                    while noSample:
                                        classUrl = elearning_domain + '/RWD/#/TMS/OnlineClass/ClassPK/' + myclass['classPK']
                                        driver.get(classUrl)
                                        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "collapse")))
                                        driver.execute_script(goToClass['onclick'])
                                        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divOnlineTest")))
                                        driver.find_element(By.XPATH, '//*[@id="divOnlineTest"]//button[1]').click()
                                        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divQuestions")))
                                        examHtml = driver.page_source
                                        lastLen = len(answerList)
                                        answerList,ansList,noSample,startNum = self.getAns(driver, qs,examHtml,answerList)
                                        #強制等待
                                        sleep(2)
                                        driver.implicitly_wait(10)
                                        allHavaAns,score = self.clickAnswerSubmit(driver,ansList,startNum,ansOneMinTime,ansOneMaxTime)
                                        beforeLen = lastLen
                                        if '100' in score or allHavaAns:
                                            noSample = False
                                        if noSample:
                                            answerList = self.getHaveAnswer(qs,driver,goToClass['onclick'])
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
    
    def getClassInfo(self, html):
        tree = etree.HTML(html)
        table = tree.xpath('//div[@class="DashBoard"]')[0]
        myclasses = table.xpath('//ul/li/div[@class="timeline-body"]')
        classInfo = []
        
        for myclass in myclasses:
            name = myclass.xpath('h2[@class="timeline-content"]/a')[0].text.strip()
            href = myclass.xpath('h2[@class="timeline-content"]/a')[0].get('href')
            try:
                score = int(myclass.xpath('td/div/span/table/tbody/tr[3]/td[2]/text()')[0].split('目前平均成績:')[-1].split(')')[0])
            except:
                score = 0
            coursePK = href.split('CoursePK/')[-1].split('/')[0]
            classPK = href.split('ClassPK/')[-1].split('/')[0]
            ispass = False #'green.gif' in myclass.xpath('td/img')[1].get('src')
            classInfo.append({'name':name,
                              'coursePK':coursePK,
                              'classPK':classPK,
                              'ispass':ispass,
                              'score':score,
                              'href': href})
        return classInfo
    
    
    def getClassBranchInfo(self,driver,myclass):
        classUrl = elearning_domain + '/RWD/#/TMS/OnlineClass/ClassPK/' + myclass['classPK']
        driver.get(classUrl)
        WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "collapse")))
        html = driver.page_source
        
        tree = etree.HTML(html)
        forceClass = False
        classes = tree.xpath('//div[@id="StudentArea"]/div[@class="row"]/div/section[contains(@class, "panel-featured-primary")]/div/div')
        ClassBranchInfo = []
        # '教材：閱讀總時數至少達30分鐘,目前進度為64分鐘。'
        mian_info = tree.xpath('//div[@id="StudentArea"]/div[@class="row"][1]/div[1]/section/div[2]/div/div[2]/div[1]/div/ul/li[2]/text()')[0] 
        # 目前分鐘
        main_progress_min = int(re.sub('\D', "", mian_info.split('分鐘')[1])) # 60
        # 總共需要分鐘
        main_total_min = int(re.sub('\D', "", mian_info.split('分鐘')[0])) # 30
        infoBox = None
        for submyclass in classes:
            ## <div class="widget-summary">
            type = submyclass.xpath('a/div/div/span')[0].get('title') # 教材 | 考試
            name = submyclass.xpath('div/div/span/a')[0].text.strip() # 課程名稱 
            onclick = submyclass.xpath('div/div/span/a')[0].get('onclick') # 
            
            # if '100' == finished.split('%')[0].replace('完成度 : ',''):
            #     finished = True
            # elif '完成度' not in finished and '未完成' not in finished and '完成' in finished and type != '教材':
            #     finished = True
            # elif '100.00' in finished:
            #     finished = True
            # else:
            #     finished = False
            # 教材:  '完成度 : 0% | 閱讀時數 : --分鐘|教材測驗成績 : 0|'
            infoBox = ''.join(submyclass.xpath('div/div/div[@class="info"]/text()')).strip() 
            finished = False
            if onclick != '': 
                if type == '教材':
                    finished = main_progress_min > main_total_min
                if type == '考試':
                    finished = '100' == re.sub('\D', "", infoBox.split('|')[0]).strip()
                    forceClass = False
                    # if forceClass:
                    #     logging.info(u'\t分鐘數不足 (已讀 {} 分鐘 , 需要 {} 分鐘) ,且已有考試,強制啟用閱讀'.format(main_progress_min,main_total_min))
                if type == '問卷':
                    finished = infoBox != '未完成'
            ClassBranchInfo.append({'type':type,
                                    'name':name,
                                    'onclick':onclick,
                                    'force':forceClass,
                                    'finished':finished,
                                    'main_progress_min':main_progress_min,
                                    'main_total_min':main_total_min,
                                    'onclick': onclick})
        return ClassBranchInfo
    
    def getNeedTolClass(self,classBranchInfo):
        skipNum = 0
        skipList = []
        lastClass = []
        beenForce = False
        for myclass in classBranchInfo:
            if myclass['onclick'] != '':
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
        
        logging.debug("into:"+classBranch['onclick'])
        driver.get_log('browser')
        loaded = False
        while not loaded:
            try:
                driver.execute_script(classBranch['onclick'])
                if driver.find_element(By.ID, 'divOrcaAlert').is_displayed():
                    driver.execute_script('setOrcaConfirmResult(true);')
                WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, "iframeCDS")))
                driver.switch_to.frame(driver.find_element(By.ID, "iframeCDS"))
                WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, "contentframe")))
                driver.switch_to.frame(driver.find_element(By.ID, "contentframe"))
                loaded = True
            except Exception as e:
                logging.debug("page is error? try again...")
        learnAll = classBranch['main_progress_min'] < classBranch['main_total_min']
        if learnAll:
            script = ''.join(open('command.txt').readlines())
        else:
            script = ''.join(open('command-com.txt').readlines())
        # logging.debug("switch_to_frame menuframe...")
        # driver.switch_to_frame(driver.find_element(By.ID, "menuframe"))
        # WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.XPATH, '//ul[@class="k-group k-treeview-lines"]')))
        html = driver.page_source
        tree = etree.HTML(html)
        
        # classes = len(tree.xpath('//ul[@class="k-group k-treeview-lines"]')[0])
        # min = (classes*6) // 60
        # sec = (classes*6) - 60*min
        # if not TrueLearn:
        #     logging.info(u"\t\t上課中...視訊量:{} 所需約:{} 分 {} 秒 ... ".format(classes,min,sec))
        # else:
        #     logging.info(u"\t\t上課中...視訊量:{}".format(classes))
        self.waitConsole(driver, 'APIAdapter initialized')
#         driver.execute_script('$(window).unbind("blur focusout");')
        learningTime,learnMin,learnSec,TotalSec = self.getLearnTime(driver,learnRatio)
        thisScript = script.format(min = str(learnMin).zfill(2),sec=str(learnSec).zfill(2),rnd=str(randint(0,99)).zfill(2)) if learnAll else script
        logging.debug('thisScript={}'.format(thisScript))
        logging.debug("exec script...:"+thisScript)
        driver.execute_script(thisScript)
        if learnAll:
            es = driver.find_elements(By.XPATH, '//div/span[@class="k-in"]')
        else:
            es = driver.find_elements(By.XPATH, '//span[@class="k-in"]/img[@src="/images/not attempted.gif"]')
            es += driver.find_elements(By.XPATH, '//span[@class="k-in"]/img[@src="/images/incomplete.gif"]')
        for e in es:
            driver.get_log('browser')
            e.click()
            self.waitConsole(driver,'APIAdapter initialized')
            
            learningTime,learnMin,learnSec,TotalSec = self.getLearnTime(driver,learnRatio)
            thisScript = script.format(min = str(learnMin).zfill(2),sec=str(learnSec).zfill(2),rnd=str(randint(0,99)).zfill(2)) if learnAll else script
            logging.debug('thisScript={}'.format(thisScript))
            logging.debug("exec script...:"+thisScript)
            driver.get_log('browser')
            driver.execute_script(thisScript)
            self.waitConsole(driver,'API.LMSSetValue(cmi.core.lesson_status, completed)')
            sleep(3)
            
        sleep(2)
        logging.debug("switch_to.default_content ...")
        driver.switch_to.default_content()
        logging.debug("switch systemframe ...")
        driver.switch_to_frame(driver.find_element(By.ID, "systemframe"))
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
        learningTime = '00:00:00'
        while learningTime == '00:00:00' or learningTime == None or '' == learningTime:
            try:
                driver.execute_script("""parent.window.API.LMSGetValue('cmi.core.total_time')""")
                learningTime = self.waitConsole(driver, 'total seconds :')
                driver.get_log('browser')
                learningTime = self.waitConsole(driver, 'total seconds :')
                driver.get_log('browser')
                if learningTime != None:
                    learningTime = learningTime.split('"')[1].replace('auto-learn:','')
                if learningTime == '00:00' or learningTime == None or '' == learningTime:
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
        examPK = re.sub('\D', "" ,url)
        answerList = []
        answerListFinal = []
        if not BigExam:
            examListUrl = elearning_domain + f"/RWD/#/Elearning/MyOnlineExam/ExamRecord/ExamPK/{examPK}/Src/"
            driver.get(examListUrl)
            sleep(2)
            WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.NAME, "ctl00")))
            text = driver.page_source
            answerUrlList = self.getAnsUrlList(text)
        else:
            examListUrl = url
            driver.get(examListUrl)
            text = driver.page_source
            answerUrlList = self.getAnsUrlList_BigExam(text)
            
        for ansurl in answerUrlList:
            driver.get(ansurl)
            WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, "divViewExamPaperBody")))
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
        trs = tree.xpath('//table[contains(@class, "sticky-enabled")]/tbody/tr')
        for tr in trs:
            td = tr.xpath('td')[-1]
            a = td.xpath('a')[0]
            lastUrl = elearning_domain + '/RWD/#/Elearning/MyOnlineExam/ExamAnswerContent/ExamPaperPK/' + re.sub('\D', '', a.get('onclick'))
            anslist.append(lastUrl)
        return anslist
    
    def getAnsUrlList_BigExam(self, html):
        anslist = []
        tree = etree.HTML(html)
        table = tree.xpath('//table[@width="100%" and @cellpadding="1"]')[0][1:]
        for e in table:
            td = e.xpath('td')[-1]
            lastUrl = etree.tostring(td).decode('utf-8').split('window.open(\'')[-1].split('\',\'\',\'')[0]
            lastUrl = lastUrl.replace('&amp;', "&")
            anslist.append(elearning_domain + '/_service/omega/pretest/' + lastUrl)
        return anslist
    
    def examParser(self, html,BigExam=False):
        mlist = []
        tree = etree.HTML(html)
        if not BigExam:
            tables = tree.xpath('//*[@id="divViewExamPaperBody"]/table')
            h5s = tree.xpath('//*[@id="divViewExamPaperBody"]/h5')
            qs = []
            chains = []
            for i, h5 in enumerate(h5s):
                chains.append((
                    h5,
                    tables[i]
                ))
                
            for target in chains:
                h5 = target[0]
                q_type = h5.xpath('text()')
                q = h5.xpath('span[1]/text()')[0][3:]
                box = target[1]
                if u'是非題' in q_type or u'單選題' in q_type or u'複選題' in q_type:
                    continue
                ansbox_trs = box.xpath('tbody/tr')
                ans = ""
                if len(ansbox_trs) <= 3:
                    # 是非題
                    for i in range(len(ansbox_trs[1:])):
                        if len(ansbox_trs[1:][i].xpath('td[2]/img')) > 0:
                            ans = "2"
                        else:
                            ans = "1"
                    firstOptionText = self.cleanText(''.join(ansbox_trs[1:][0].xpath('td[3]/text()')))
    #                 list.append(ans+"\t\t"+q)
                else:
                    # 選擇題 or 複選題
                    for i, tr in enumerate(ansbox_trs):
                        if len(tr.xpath('td[2]/i')) > 0:
                            if type(ans) is str:
                                if ans == "": #沒有答案
                                    ans = str(i + 1)  # +','+ansbox[1:][i].xpath('td[3]')[0].xpath('normalize-space()')
                                    firstOptionText = self.cleanText(''.join(tr.xpath('td[3]/text()')).strip())
                                else: # 已經有答案  代表複選題
                                    ans = [ans,str(i + 1)]
                            elif type(ans) is list: #複選題
                                ans.append(str(i + 1))
    #             print firstOptionText.encode('big5','ignore')
                mlist.append((ans, q, firstOptionText))
        else:
            table = tree.xpath('/html/body/table/tbody/tr[2]/td/table[@class="table" and @width="100%" and not(@cellpadding="3")]/tbody/tr')
            for box in table:
                if len(box.xpath('td[@class="thead"]')) > 0:
                    continue
                q_type = box.xpath('td[@class="tdrowbody"]/table[@class="tdrowbody"]/tbody/tr/td[@class="topicCaption"]')[0].xpath('normalize-space()')
                q_type = ''.join(q_type.split('.')[1:]).split(u' (題目配分：')[0]
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
                mlist.append((ansText, self.cleanText(q_type.strip())))
        return mlist
    
    
    def getAns(self,driver, qs,examHtml,answerList,BigExam=False):
        tree = etree.HTML(examHtml)
        table = tree.xpath('//*[@id="divQuestions"]/div[@class="examBlock"]')
        no = 0
        ansList = []
        noSample = False
        driver.execute_script('$(window).unbind("blur focusout");')
        if not BigExam:
            for i,examqution in enumerate(table):
                no += 1
                examqutionText = (''.join(examqution.xpath('//div[1]/span/text()')).strip())[3:]
                # examFirstOptionText = ''.join(examqution.xpath('div/table/tbody/tr[1]/td/label/text()')).strip()
                # id = re.sub('.*(\d+).*',r'\1',examqution.xpath('span')[0].get('id')[0:-2]).zfill(2)
                # if i == 0:
                #     startNum = id
                find = False
                for ans, q, firstOptionText in answerList:
                    if self.cleanText(examqutionText) in q:
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
#                     el = driver.find_element(By.ID, 'QuestionItemList_ctl{}_isAnswer_{}'.format(str(int(startNum)+i).zfill(2),int(a)-1))
#                     action = webdriver.common.action_chains.ActionChains(driver)
#                     action.move_to_element_with_offset(el,6,6)
#                     action.click()
#                     action.perform()
        try:
            driver.maximize_window()
        except:
            logging.debug('fail to max ,skip:'+traceback.format_exc())
        driver.find_element(By.NAME, "saveBtn").click()
        if hide :
            try:
                driver.set_window_position(-10000,0)
            except:
                logging.debug('fail to hide , skip:'+traceback.format_exc())
#         saveBtn = driver.find_element(By.NAME, "saveBtn")
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
        score = driver.find_element(By.XPATH, '//span[@id="Score"]').text.split('分')[0]
        logging.info(u'\t分數:{}'.format(score))
        
        return allHavaAns,score
    
    def questionnaire(self,driver,qs,myclass):
        PK = myclass['onclick'][myclass['onclick'].find('PK='):]
        PK = PK[:PK.find('&')]
        PK = PK.replace('PK=','')
        self.goToHome(driver)
        html = driver.page_source
        userPK = html[html.find('UserPK='):]
        userPK = userPK[:userPK.find('&')]
        userPK = userPK.replace('UserPK=',"")
        driver.get(elearning_domain + '/_service/omega/survey2/DoClassSurveyforMobile.asp?PK={}'.format(PK))
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
        rs = qs.post(elearning_domain + '/_service/omega/survey2/ActionSurveyScoreMobile.asp',data = data)
        
    
        
    
                
    
    
    
    
    
    def loggin(self,driver,ac,pw,hide = True):
        qs = requests.Session()
        qs.verify = False
        while True:
            try:
                web = u'數位學習平台'
                logging.info(u'登入中...')
                driver.get(elearning_domain + '/WcmsModules/Portal/FullFrameLogin/index.aspx')
                sleep(0.5)
                driver
                driver.find_elements(By.ID, "UserID")[0].send_keys(ac)
                driver.find_elements(By.ID, "Password")[0].send_keys(pw)
                    
                if hide :
                    driver.set_window_position(-10000,0)
                driver.find_elements(By.ID, "btnSignIn")[0].click()
                for cookie in driver.get_cookies():
                    qs.cookies.set(cookie['name'], cookie['value'])
                logo = driver.find_element(By.CLASS_NAME, 'logo-container')
                if logo == None:
                    raise Exception('登入失敗!')
                WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.CLASS_NAME, "DashBoard")))
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
        driver.get(elearning_domain + '/Index.aspx')
        driver.switch_to_frame(driver.find_element(By.NAME, "bottomframe"))
        driver.switch_to_frame(driver.find_element(By.NAME, "display"))
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
        logging.getLogger("requests").setLevel(logging.ERROR)
        logging.getLogger("selenium").setLevel(logging.ERROR)
class detection_event(Thread):
    def run(self):
        while True:
            if autoElearning.detection:
                driver = autoElearning.driver
                logging = autoElearning.loggin
                try:
                    driver.switch_to.default_content()
                    btn = WebDriverWait(driver,3).until(EC.presence_of_element_located((By.XPATH, '//input[@onclick="fnTimeIsUpCheck()"]')))
                    if btn.is_displayed():
                        logging.info(u'\t點擊繼續閱讀...')
                        btn.click()
                    driver.switch_to_frame(driver.find_element(By.ID, "menuframe"))
                except Exception as e:
                    logging.error(e)
                sleep(0.1)
detection_event().start()
autoElearning().run()