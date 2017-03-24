from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
import json
import os
import re
import xlwt
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary


def ReadConfig():
    file = 'tt.config'
    fp = open(file, 'r',encoding= 'utf-8')
    config = json.load(fp)
    fp.close()
    return config


def DoBaidu():
    url = 'https://www.baidu.com/'
    driver.get(url)
    kw = '日历'
    driver.find_element_by_id("kw").send_keys(kw)
    driver.find_element_by_id('su').submit()
    time.sleep(2)
    driver.get_screenshot_as_file('Step1日历.png')
    driver.find_element_by_xpath('//*[@id="1"]/div[1]/div[1]/div[1]/div[1]/div[2]/div/div[1]/div[2]/div').click()
    driver.find_element_by_xpath('//*[@id="1"]/div[1]/div[1]/div[1]/div[1]/div[2]/div/div[2]/div/ul/li[2]').click()
    driver.get_screenshot_as_file('Step2切换.png')
    data = driver.find_element_by_xpath('//*[@id="1"]/div[1]/div[1]/div[2]/p[3]/span[3]').text
    print(data)
    tdNode = driver.find_element_by_xpath('//*[@id="1"]/div[1]/div[1]/div[1]/div[2]/table/tbody/tr[1]')
    thNode = tdNode.find_element_by_xpath('./th[5]')
    print(thNode.text)
    print('测试成功！')

def printLine(output):
    print(output)
def ExcProvinceCheck():
    url = 'http://mz.dmw.gov.cn/index3.php'
    driver.get(url)
    sel = driver.find_element_by_xpath('//*[@id="province"]').find_elements_by_tag_name('option')
    num = len(sel)
    count = 0
    for ind in range(0, num):
        each = sel[ind]
        each.click()
        time.sleep(0.5)
        count += 1
        driver.get_screenshot_as_file(each.text + '.png')
        if (count > 3):
            break
    sel[10].click()
    driver.get_screenshot_as_file('Step1江苏.png')
    province = Select(driver.find_element_by_xpath('//*[@id="province"]'))
    count = 0
    for each in province.options:
        if (count > 3):
            break
        count += 1
        print('sel ' + each.text)
    sx = province.select_by_value('30')
    time.sleep(1)
    driver.get_screenshot_as_file('Step1陕西.png')
    city = Select(driver.find_element_by_xpath('//*[@id="city"]'))
    bj = city.select_by_value('313')
    time.sleep(1)
    driver.get_screenshot_as_file('Step2宝鸡.png')
    print('测试成功！')

def WaitControlById(str):
    loginWait = float(config['loginTime'])
    while 1:
        try:
            time.sleep(loginWait)
            print(str+'元素 正在加载，请稍等...')
            driver.find_element_by_id(str)
        except:
            continue
        else:
            # driver.find_element_by_name(str).get_attribute('innerHTML')
            # driver.find_element_by_id(str).click()
            # print('加载完成，正在切换菜单...')
            return

def WaitFrameByClass(str):
    loginWait = float(config['loginTime'])
    while 1:
        try:
            time.sleep(loginWait)
            print('正在加载，请稍等...')
            driver.find_element_by_class_name(str)
        except:
            continue
        else:
            #driver.find_element_by_name(str).get_attribute('innerHTML')
            #driver.find_element_by_id(str).click()
            #print('加载完成，正在切换菜单...')
            return

def WaitControlClickByName(str):
    loginWait = float(config['loginTime'])
    while 1:
        try:
            time.sleep(loginWait)
            print( '正在加载，请稍等...')
            driver.find_element_by_name(str)
        except:
            continue
        else:
            #driver.find_element_by_name(str).get_attribute('innerHTML')
            driver.find_element_by_name(str).click()
            print(driver.find_element_by_name(str).get_attribute('innerHTML') +' 元素加载完成，正在切换菜单...')
            return

def InitToMenu():
    userName=config['userName']
    pwd = config['password']

    loginWait = float(config['loginTime'])
    driver.get(config['url'])
    driver.get_screenshot_as_file('login.png')
    driver.find_element_by_id('j_username').send_keys(userName)
    driver.find_element_by_id('j_password').send_keys(pwd)

    driver.find_element_by_id('btnlogin').click()
    time.sleep(loginWait)
    WaitControlClickByName('工单管理')
    WaitControlClickByName('工单管理业务开通')
    WaitControlClickByName('工单管理业务开通业务开通工单')

def FilterSearch():
    # WaitControlClickById('highSearchBtn1')
    WaitFrameByClass('iframe1')
    driver.switch_to.frame(driver.find_element_by_class_name("iframe1"))
    # driver.find_element_by_id('highSearchBtn1').click()
    WaitControlById('highSearchBtn1')

    # time.sleep(20)
    driver.get_screenshot_as_file('Menu.png')
    driver.execute_script('console.log(\'1\')')
    driver.execute_script('showHighSearchDiv()')
    # return  driver
    driver.find_element_by_name('undefined').click()
    driver.find_element_by_class_name('l-box-select-inner').find_element_by_xpath('table/tbody/tr[1]/td').click()
    driver.find_element_by_xpath('//*[@id="highSearchDivForm"]/ul/li[1]/div/span').click()
    #time.sleep(8)
    #driver.find_element_by_xpath('//*[@id="highSearchDivForm"]/ul/li[3]/div').click()

def OrderRequest(i):
    driver.execute_script('getRowData('+str(i)+')')
    time.sleep(waitTime)

    if "铁通转网" not in driver.find_element_by_id('crmRemark').get_attribute('value'):
        return

    driver.find_element_by_xpath('//*[@id="workOrderDetailsTitle"]/ul/li[2]/a').click()
    time.sleep(waitTime)
    if "前台预约" in driver.find_element_by_xpath('//*[@id="historyLinkInfoDiv|2|r1001|c103"]/div').get_attribute('innerHTML') and "成功" in driver.find_element_by_xpath('//*[@id="historyLinkInfoDiv|2|r1001|c107"]/div').get_attribute('innerHTML'):
        return
    printLine('该用户需要铁通转网！')

    driver.execute_script('showInstruPreDiv()')
    time.sleep(waitTime)


    WaitControlClickByName('userHandlerName')
    driver.find_element_by_name('userHandlerName').send_keys(assignWorker)
    time.sleep(10)
    #To do:assign worker automatically by counts
    workerList= driver.find_elements_by_xpath('//div[@class="l-box-select-inner"]')[3]
    print(workerList.find_element_by_tag_name('td').get_attribute('text'))
    workerList.find_element_by_tag_name('td').click()

    #workers={}
    #for each in workerList:
    #driver.find_element_by_xpath('//div[@class="l-box-select-inner"]/table/tbody/tr[5]/td').click()

    #click calendar
    driver.find_elements_by_class_name('l-trigger-icon')[12].click()
    time.sleep(3)
    driver.find_elements_by_xpath('//div[@class="l-box-dateeditor-body"]/table/tbody/tr[5]/td[5]')[2].click()
    #time.sleep(3)
    #driver.find_element_by_id('remark').send_keys("铁通转网啦小蕾蕾")
    assignCount[0] =assignCount[0]-1
    try:
        #driver.execute_script('alert()')
        driver.execute_script('preSucBtn()')
    except:
        driver.switch_to.alert.accept()
        print("员工： "+assignWorker+"已派单"+str(i)+"次， 剩余 "+str(assignCount[0])+"次")
        driver.get_screenshot_as_file('assigned.png')
        time.sleep(3)

def Exctt():
    InitToMenu()
    FilterSearch()
    for i in range(0,30):
        if(assignCount[0]<=0):
            return
        OrderRequest(i)
    print("运行结束！")
    return


if __name__ == '__main__':
    config = ReadConfig()
    waitTime = float(config['waitTime'])


    #driver = webdriver.Firefox(executable_path="./geckodriver.exe")

    #无浏览器
    #driver = webdriver.PhantomJS(executable_path="./phantomjs.exe")
    #chromedriver = "./chromedriver.exe"
    #os.environ["webdriver.chrome.driver"] = chromedriver
    #driver = webdriver.Chrome(chromedriver)
    #driver = webdriver.Chrome(executable_path="./chromedriver.exe")
    mode = config['mode']
    assignCount=[]
    assignCount.append(int(config['assignCount']))
    assignWorker = config['assignWorker']
    if (mode == 'baidu'):
        DoBaidu()

    elif (mode == 'area'):
        ExcProvinceCheck()
    elif (mode == 'standardTT'):
        #无浏览器
        driver = webdriver.PhantomJS(executable_path="./phantomjs.exe")
        Exctt()
    else:
        #启动FireFox，可视化界面
        binary = FirefoxBinary(r'D:\Program Files (x86)\Mozilla Firefox\firefox.exe')
        driver = webdriver.Firefox(firefox_binary=binary)
        Exctt()
    #os.system("pause")
