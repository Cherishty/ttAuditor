from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
import json
import os
import re
import codecs
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import traceback

global assignWorker
global assignIndex
global assignableNums
assignableNums = 1


def ReadConfig():
    file = 'tt.config'
    fp = open(file, 'r', encoding='utf-8')
    config = json.load(codecs.open('tt.config', 'r', 'utf-8-sig'))
    fp.close()
    return config


def printLine(output):
    print(output)
    file_object.write(output+'\n')

def WaitControlById(str):
    while 1:
        try:
            time.sleep(loginWait)
            printLine(str + '元素 正在加载，请稍等...')
            driver.find_element_by_id(str).get_attribute('id')
        except:
            continue
        else:
            # driver.find_element_by_name(str).get_attribute('innerHTML')
            # driver.find_element_by_id(str).click()
            # printLine('加载完成，正在切换菜单...')
            return


def SwitchFrameByClass(str):
    while 1:
        try:
            time.sleep(loginWait)
            printLine(str + ' 正在加载，请稍等...')
            driver.find_element_by_class_name(str).get_attribute('class')
        except:
            continue
        else:
            driver.switch_to.frame(driver.find_element_by_class_name("iframe1"))
            return


def WaitControlClickByName(str, speed=0):
    while 1:
        try:
            if (speed == 0):
                time.sleep(loginWait)
            else:
                time.sleep(waitTime)
            printLine(str + ' 正在加载，请稍等...')
            # driver.get_screenshot_as_file(str + '.png')
            driver.find_element_by_name(str)
        except:
            continue
        else:
            # driver.find_element_by_name(str).get_attribute('innerHTML')
            driver.find_element_by_name(str).click()
            printLine(driver.find_element_by_name(str).get_attribute('name') + ' 元素加载完成，正在切换菜单...')
            return


def InitToMenu():
    userName = config['userName']
    pwd = config['password']

    loginWait = float(config['loginTime'])
    driver.get(config['url'])
    driver.get_screenshot_as_file('picture/Step1_login.png')
    driver.find_element_by_id('j_username').send_keys(userName)
    driver.find_element_by_id('j_password').send_keys(pwd)

    driver.find_element_by_id('btnlogin').click()
    time.sleep(loginWait)
    WaitControlClickByName('工单管理')
    time.sleep(loginWait)
    WaitControlClickByName('工单管理业务开通')
    WaitControlClickByName('工单管理业务开通业务开通工单')


def FilterSearch():
    WaitControlById('highSearchBtn1')
    driver.get_screenshot_as_file('picture/Step2_Menu.png')
    driver.execute_script('showHighSearchDiv()')
    driver.find_element_by_name('undefined').click()
    if(config['behavior']=='order'):
        driver.find_element_by_class_name('l-box-select-inner').find_element_by_xpath('table/tbody/tr[1]/td').click()
    if(config['behavior']=='auditor'):
        driver.find_element_by_class_name('l-box-select-inner').find_element_by_xpath('table/tbody/tr[4]/td').click()
    driver.find_element_by_xpath('//*[@id="highSearchDivForm"]/ul/li[1]/div/span').click()
    time.sleep(3)


'''
def AssignWorkerHalf():
    driver.find_element_by_name('userHandlerName').send_keys(assignWorker)
    time.sleep(10)
    driver.get_screenshot_as_file('picture/Step4_待安排人员.png')
    # To do:assign worker automatically by counts
    workerList = driver.find_elements_by_xpath('//div[@class="l-box-select-inner"]')[3]
    printLine(workerList.find_element_by_tag_name('td').get_attribute('text'))
    workerList.find_element_by_tag_name('td').click()

    # workers={}
    # for each in workerList:
    # driver.find_element_by_xpath('//div[@class="l-box-select-inner"]/table/tbody/tr[5]/td').click()
'''


def GetAssignWorker():
    global assignableNums
    global assignIndex
    workerList = driver.find_elements_by_xpath('//div[@class="l-box-select-inner"]')[3]
    if totalCount[0] == 0:
        assignIndex = 0
        assignableNums = len(workerList.find_elements_by_tag_name('td'))
        i = 0
        for each in workerList.find_elements_by_tag_name('td'):
            assignedInfo = each.get_attribute("innerHTML")
            assignedName = re.findall('(.*?)    待处理工单数', assignedInfo, re.S)[0]
            assignableWorker[assignedName] = 0
            assignableMap.append(assignedName)
        printLine("可安排的人员共有： " + str(assignableNums) + "人")
    if (assignIndex >= assignableNums):
        assignIndex = 0
    assignIndex += 1
    return assignableMap[assignIndex - 1]


def AssignWorker():
    WaitControlClickByName('userHandlerName', 1)
    global assignWorker
    if 0 == assignMode:
        assignWorker = GetAssignWorker()
    driver.find_element_by_name('userHandlerName').send_keys(assignWorker)
    time.sleep(10)
    driver.get_screenshot_as_file('picture/Step4_待安排人员.png')
    workerList = driver.find_elements_by_xpath('//div[@class="l-box-select-inner"]')[3]
    printLine(workerList.find_element_by_tag_name('td').get_attribute('text'))
    workerList.find_element_by_tag_name('td').click()


def ValidateOrderStatus(i):
    try:
        driver.execute_script('getRowData(' + str(i) + ')')
    except:
        return -1
    time.sleep(loginWait)
    if (debugMode < 2):
        if "铁通转网" not in driver.find_element_by_id('crmRemark').get_attribute('value'):

            return 0

    driver.get_screenshot_as_file('picture/Step3.5_铁通转网.png')
    driver.find_element_by_xpath('//*[@id="workOrderDetailsTitle"]/ul/li[2]/a').click()
    time.sleep(waitTime)
    if (debugMode < 2 and config['behavior']=='order'):
        lastStatus = driver.find_element_by_xpath(
            '//*[@id="historyLinkInfoDivgrid"]/div[4]/div[2]/div/table/tbody/tr[last()]')
        currentProcess = lastStatus.find_element_by_xpath('./td[3]/div').get_attribute('innerHTML')
        currentStatus = lastStatus.find_element_by_xpath('./td[7]/div').get_attribute('innerHTML')
        if "前台预约" in currentProcess and "处理中" in currentStatus:
            return 1
        else:
            return 0
    return 1


def OrderRequest():
    # 安排人员
    AssignWorker()

    # 预约时间
    ClickCalendar()


    assignedInfo = driver.find_element_by_name('userHandlerName').get_attribute("value")
    assignedName = re.findall('(.*?)    待处理工单数', assignedInfo, re.S)[0]
    global assignWorker
    if (assignWorker == assignedName):
        assignableWorker[assignedName] += 1
    else:
        assignWorker = assignedName
        assignableWorker[assignedName] = 1
    printLine("员工： " + assignedName + "已派单" + str(assignableWorker[assignedName]) + "次， 剩余 " + str(
            assignCount[0] - assignableWorker[assignedName]) + "次")


def ProcessRequest(i):
    if i==29:
        pass
    isValid = ValidateOrderStatus(i)
    if isValid != 1:
        printLine('该用户不需要铁通转网！')
        driver.get_screenshot_as_file('picture/Step3.7_不转网.png')
        return isValid
    number=driver.find_element_by_name('pbossOrderCode').get_attribute('value')
    orderUser=driver.find_element_by_name('handlerUser').get_attribute('value')
    printLine('正在'+config["behavior"]+'该工单： '+number+ orderUser)

    if config['behavior']=='order':
        driver.execute_script('showInstruPreDiv()')
        OrderRequest()
    driver.get_screenshot_as_file('picture/Step5_处理完毕.png')
    totalCount[0] += 1
    try:
        if (debugMode > 0):
            driver.execute_script('alert()')

        elif (debugMode == 0):
            if config['behavior']=='order':
                driver.execute_script('preSucBtn()')
            elif config['behavior']=='auditor':
                driver.execute_script('showCheckDaiWeiDiv()')
                driver.find_element_by_id('checkDaiweiDivTool').find_element_by_xpath('./a[1]').click()

    except:
        driver.get_screenshot_as_file('picture/Step6_popup.png')
        if mode == 'amazingTT':
            driver.switch_to.alert.accept()
        time.sleep(waitTime)
        if (debugMode >0):
            return 0
        return 1

def ClickCalendar():
    driver.find_elements_by_class_name('l-trigger-icon')[12].click()
    time.sleep(waitTime)
    nextMonth = driver.find_elements_by_xpath('//div[@class="l-box-dateeditor-header"]')[2]
    nextMonth.find_element_by_xpath('./div[4]/span').click()

    driver.find_elements_by_xpath('//div[@class="l-box-dateeditor-body"]/table/tbody/tr[6]/td[1]')[2].click()


def NavToNextPage(hasAssigned):
    if debugMode>0:
        driver.get_screenshot_as_file('picture/Step5.5_准备换页前.png')
    if not hasAssigned or hasAssigned == 0:
        # 关闭预约界面
        driver.execute_script('closeWorkOrderDetailsDiv()')

    if (mode == 'amazingTT'):
        #滚动到最下以防无‘下一页’按键
        driver.execute_script('document.documentElement.scrollTop=20')
    if debugMode>0:
        driver.get_screenshot_as_file('picture/Step6_准备换页.png')
    nextpage = driver.find_element_by_xpath('//*[@id="businessListGrid"]/div[5]/div/div[8]/div[1]/span')
    nextpage.click()
    time.sleep(3)


def Exctt():
    InitToMenu()
    SwitchFrameByClass('iframe1')
    if (debugMode < 2):
        FilterSearch()
    currentPage = 1
    totalPage = int(
            driver.find_element_by_xpath('//*[@id="businessListGrid"]/div[5]/div/div[6]/span/span').get_attribute(
                    'innerHTML'))

    while currentPage <= totalPage:
        printLine('当前在第' + str(currentPage) + '页， 共计' + str(totalPage) + '页')
        driver.get_screenshot_as_file('picture/Step3_当前页.png')
        hasAssigned = 0
        item = 1
        if (debugMode < 3):
            item = 30
        for i in range(0, item):
            global assignableNums
            if (config['behavior']=='order' and ((assignMode and assignCount[0] <= assignableWorker[assignWorker]) or totalCount[0] >= assignableNums *
                assignCount[0])):
                return
            hasAssigned = ProcessRequest(i)
            if hasAssigned == -1:
                printLine('运行结束')
                return 0

        NavToNextPage(hasAssigned)
        currentPage += 1
        totalPage = int(
            driver.find_element_by_xpath('//*[@id="businessListGrid"]/div[5]/div/div[6]/span/span').get_attribute(
                'innerHTML'))
    return


if __name__ == '__main__':
    file_object = open('runTT.log', 'w+')
    printLine('Welcome ttAuditor v6.6')
    config = ReadConfig()
    waitTime = float(config['waitTime'])
    loginWait = float(config['loginTime'])
    mode = config['mode']
    debugMode=int(config['debugMode'])
    assignMode = 1
    assignCount = []
    if (config['assignMode']['useMode'] == 'fullAuto'):
        assignMode = 0
        assignCount.append(int(config['assignMode']['fullAuto']['assignCount']))
    elif (config['assignMode']['useMode'] == 'halfAuto'):
        assignMode = 1
        assignCount.append(int(config['assignMode']['halfAuto']['assignCount']))
    # chromedriver = "./chromedriver.exe"
    # os.environ["webdriver.chrome.driver"] = chromedriver
    # driver = webdriver.Chrome(chromedriver)
    # driver = webdriver.Chrome(executable_path="./chromedriver.exe")
    global assignWorker
    assignWorker = config['assignMode']['halfAuto']['assignWorker']
    assignableWorker = {}
    assignableMap = []
    totalCount = []
    totalCount.append(0)
    assignableWorker[assignWorker] = 0
    printLine("behavior:"+config["behavior"] )
    if (config["behavior"] == "order"):
        printLine("assignMode: " + config['assignMode']['useMode'])
    try:

        if (mode == 'standardTT'):
            # 无浏览器
            driver = webdriver.PhantomJS(executable_path="./phantomjs.exe")
            Exctt()
        elif (mode == 'amazingTT'):
            # 启动FireFox，可视化界面
            binary = FirefoxBinary(r'D:\Program Files (x86)\Mozilla Firefox\firefox.exe')
            driver = webdriver.Firefox(firefox_binary=binary)
            Exctt()
        else:
            printLine('未能识别运行模式，请检查config')
            #chromedriver = "./chromedriver.exe"
            #os.environ["webdriver.chrome.driver"] = chromedriver
            #driver = webdriver.Chrome(chromedriver)
            #driver = webdriver.Chrome(executable_path="chromedriver.exe")
            #Exctt()
    except Exception as e:
        printLine('traceback.format_exc():\n%s' % traceback.format_exc())
    finally:
        printLine('共处理工单数：' +str(totalCount[0]))
        printLine("运行结束！")
        file_object.close()
        os.system("pause")
