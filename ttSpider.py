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
    fp = open(file, 'r')
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
    time.sleep(8)
    driver.find_element_by_xpath('//*[@id="highSearchDivForm"]/ul/li[3]/div').click()

def OrderRequest(i):
    driver.execute_script('getRowData('+str(i)+')')
    time.sleep(waitTime)
    if "铁通转网" not in driver.find_element_by_id('crmRemark').get_attribute('innerHTML'):
        return
    driver.find_element_by_xpath('//*[@id="workOrderDetailsTitle"]/ul/li[2]/a').click()
    time.sleep(waitTime)
    if "前台预约" in driver.find_element_by_xpath('//*[@id="historyLinkInfoDiv|2|r1001|c103"]/div').get_attribute('innerHTML') and "成功" not in driver.find_element_by_xpath('//*[@id="historyLinkInfoDiv|2|r1001|c107"]/div'):
        return
    driver.execute_script('alert(\'1\')')


def Exctt():
    InitToMenu()
    FilterSearch()
    for i in range(0,30):
        OrderRequest(i)
    driver.get_screenshot_as_file('Menu1.png')

    return

def Dott():
    try:
        temp = driver.find_element_by_id('DropDownList2')
    except:
        print('登录失败，未定位到成功页面')
        driver.get_screenshot_as_file('loginFail.png')

    driver.get_screenshot_as_file('loginSucc.png')

    num1 = len(driver.find_element_by_id('DropDownList2').find_elements_by_tag_name('option'))
    for ind1 in range(0, num1):
        if(ind1!=2 and ind1 != 4):
            continue
        option1 = driver.find_element_by_id('DropDownList2').find_elements_by_tag_name('option')[ind1]
        depart = option1.text
        print("切换到1级菜单" + depart)
        option1.click()
        time.sleep(waitTime)
        result = {}

        num2 = len(driver.find_element_by_id('DropDownList5').find_elements_by_tag_name('option'))

        try:
            for ind2 in range(0, num2):
                option2 = driver.find_element_by_id('DropDownList5').find_elements_by_tag_name('option')[ind2]
                dis = option2.text
                print("切换到2级菜单" + dis)
                option2.click()
                time.sleep(waitTime)

                num3 = len(driver.find_element_by_id('DropDownList3').find_elements_by_tag_name('option'))
                for ind3 in range(0, num3):
                    option3 = driver.find_element_by_id('DropDownList3').find_elements_by_tag_name('option')[ind3]
                    area = option3.text
                    print("切换到3级菜单" + area)
                    option3.click()
                    time.sleep(waitTime)

                    num4 = len(driver.find_element_by_id('DropDownList4').find_elements_by_tag_name('option'))
                    for ind4 in range(0, num4):
                        try:
                            #print("切换下拉菜单ing")
                            option4 = driver.find_element_by_id('DropDownList4').find_elements_by_tag_name('option')[ind4]
                            unit = option4.text
                            print("切换到4级菜单" + unit)
                            option4.click()
                            time.sleep(waitTime)
                        except:
                            print("切换异常")
                            driver.get_screenshot_as_file('switchERROR.png')
                        driver.get_screenshot_as_file('dropdownFound.png')
                        info = []

                        info.append(depart)
                        info.append(dis)
                        info.append(area)
                        info.append(unit)
                        tbody = driver.find_element_by_tag_name('tbody')
                        items = tbody.find_elements_by_tag_name('tr')

                        count = 0
                        max = 0
                        for item in items:
                            if (count == 0):
                                count = 1
                                continue

                            tditem = item.find_element_by_xpath('./td[11]').text
                            if (not tditem or tditem == ' '):
                                break
                            port = tditem.split('/')[-1]
                            #print("当前端口为：" + port)
                            if (int(port) > max):
                                max = int(port)

                        info.append(str(max))
                        print(unit + '最大端口为: ' + str(max))
                        if (dis not in result.keys()):
                            result[dis] = []
                        result[dis].append(info)

        finally:
            file = xlwt.Workbook()
            table = file.add_sheet('首页', cell_overwrite_ok=True)
            table.write(0, 0, '测试页')
            for dis in result.keys():
                i = 0
                print("正在保存 " + dis)
                table = file.add_sheet(dis, cell_overwrite_ok=True)
                for line in result[dis]:
                    j = 0
                    for each in line:
                        table.write(i, j, each)
                        j += 1
                    i += 1
            fielname = depart + '.xls'
            file.save(fielname)


if __name__ == '__main__':
    config = ReadConfig()
    waitTime = float(config['waitTime'])
    #启动FireFox，可视化界面
    binary = FirefoxBinary(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe')
    driver = webdriver.Firefox(firefox_binary=binary)

    #driver = webdriver.Firefox(executable_path="./geckodriver.exe")

    #无浏览器
    #driver = webdriver.PhantomJS(executable_path="./phantomjs.exe")
    #chromedriver = "./chromedriver.exe"
    #os.environ["webdriver.chrome.driver"] = chromedriver
    #driver = webdriver.Chrome(chromedriver)
    #driver = webdriver.Chrome(executable_path="./chromedriver.exe")
    mode = config['mode']
    if (mode == 'baidu'):
        DoBaidu()

    elif (mode == 'area'):
        ExcProvinceCheck()
    else:
        Exctt()
    #os.system("pause")
