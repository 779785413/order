from selenium import webdriver
import time
import traceback
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import *
import threading
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import  inspect
import ctypes
TEST_ORDER = 6060931767170302
PRINTSWITCH = 1
BTNENABLE = 1
procThread = None
statlt = ['AL','AK','AS','AZ','AR','AE','AA','AP','CA','CO','CT','DE','DC','FL','GA','GU','HI','ID','IL','IN','IA','KS','KY','LA','ME','MP','MH','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','PR','RI','SC','SD','TN','TX','UT','VT','VI','VA','WA','WV','WI','WY']
#停止线程
def stopThread(tid,cmd):
    tid = ctypes.c_long(tid)
    print('tid is : ',tid)
    if not inspect.isclass(cmd):
        cmd = type(cmd)
        print('cmd is : ', cmd)
    print('ret is ',ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(cmd)))
    button['state'] = 'normal'
def switchToWindwos(windowName,driver):
    windows = driver.window_handles
    for SingleWindow in windows:
        driver.switch_to.window(SingleWindow)
        if driver.title == windowName:
            break
    return driver

def webQuteOperate(webQuteDriver,addrInfoDir,shopInfo,order):
    webQuteDriver.find_element_by_id('searchValue').send_keys("124858")
    webQuteDriver.find_element_by_id('search').click()
    #等待页面加载完成
    while (1):
        try:
            webQuteDriver.switch_to.frame("webquote_customer0Frame")
            webQuteDriver.switch_to.frame("cust_detail_shippingFrame")
            if webQuteDriver.find_element_by_xpath('//*[@id="tdBizName"]/input[2]').is_displayed() == True:
                break
            else:
                webQuteDriver.switch_to_default_content()
        except:
            webQuteDriver.switch_to.parent_frame()

        try:
            webQuteDriver.switch_to.frame("webquote_customer1Frame")
            webQuteDriver.switch_to.frame("cust_detail_shippingFrame")
            if webQuteDriver.find_element_by_xpath('//*[@id="tdBizName"]/input[2]').is_displayed() == True:
                break
            else:
                webQuteDriver.switch_to_default_content()
        except:
            webQuteDriver.switch_to.parent_frame()

        try:
            webQuteDriver.switch_to.frame("webquote_customer2Frame")
            webQuteDriver.switch_to.frame("cust_detail_shippingFrame")
            if webQuteDriver.find_element_by_xpath('//*[@id="tdBizName"]/input[2]').is_displayed() == True:
                break
            else:
                webQuteDriver.switch_to_default_content()
        except:
            webQuteDriver.switch_to.parent_frame()

        try:
            webQuteDriver.switch_to.frame("webquote_customer3Frame")
            webQuteDriver.switch_to.frame("cust_detail_shippingFrame")
            if webQuteDriver.find_element_by_xpath('//*[@id="tdBizName"]/input[2]').is_displayed() == True:
                break
            else:
                webQuteDriver.switch_to_default_content()
        except:
            webQuteDriver.switch_to.parent_frame()
        time.sleep(1)

    #等待页面加载完成
    WebDriverWait(webQuteDriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tdBizName"]/input[2]')))
    WebDriverWait(webQuteDriver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="tdBizName"]/input[2]')))

    webQuteDriver.find_element_by_xpath('//*[@id="tdBizName"]/input[2]').click()

    # 等待页面加载完成
    WebDriverWait(webQuteDriver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="locName"]')))
    WebDriverWait(webQuteDriver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="locName"]')))

    time.sleep(0.5)
    webQuteDriver.find_element_by_xpath('//*[@id="locName"]').send_keys(addrInfoDir['company'])

    webQuteDriver.find_element_by_xpath('//*[@id="addr"]').send_keys(addrInfoDir['addr'])
    #如果大地址超过30个则会报错，报错接受警告窗口
    if len(addrInfoDir['addr']) >= 30:
        webQuteDriver.find_element_by_xpath('//*[@id="addr2"]').click()
        webQuteDriver.switch_to_alert().accept()

        #弹窗之后等待元素可见
        WebDriverWait(webQuteDriver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="addr2"]')))
        #警告窗口接受后只能硬等待2s，才能填充之后得参数

    #字典长度为9说明有小地址，则填充小地址，反之不用填写
    if len(addrInfoDir) == 9:
        webQuteDriver.find_element_by_xpath('//*[@id="addr2"]').click()
        webQuteDriver.find_element_by_xpath('//*[@id="addr2"]').send_keys(addrInfoDir['addr2'])
        #如果小地址查过30个字符报错，接受警告窗口
        if len(addrInfoDir['addr2']) >= 30:
            webQuteDriver.find_element_by_xpath('//*[@id="city"]').click()
            webQuteDriver.switch_to_alert().accept()

            WebDriverWait(webQuteDriver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="city"]')))
            webQuteDriver.find_element_by_xpath('//*[@id="city"]').click()
            #警告窗口接受后只能硬等待2s，才能填充之后得参数
    #城市、邮编、州混合在一个字符串中需要拆分如：ANN ARBOR, MI 48109-1432
    cityInfoList = addrInfoDir['city'].split(',')
    city = cityInfoList[0]
    stat = cityInfoList[1].split(' ')[1]
    post = cityInfoList[1].split(' ')[2]
    post = post.split('-')[0]

    webQuteDriver.find_element_by_xpath('//*[@id="city"]').send_keys(city)
    webQuteDriver.find_element_by_xpath('//*[@id="zipCode"]').send_keys(post)

    webQuteDriver.find_element_by_xpath('//*[@id="contactName"]').send_keys(addrInfoDir['user'])
    webQuteDriver.find_element_by_xpath('//*[@id="phone"]').send_keys(addrInfoDir['phone'])
    webQuteDriver.find_element_by_xpath('//*[@id="email"]').send_keys(addrInfoDir['email'])
    webQuteDriver.find_element_by_xpath('//*[@id="euPO"]').send_keys(addrInfoDir['custordr'])

    #填充州
    statElement = webQuteDriver.find_element_by_xpath('//*[@id="state"]')
    statElement.click()
    stats = statElement.find_elements_by_tag_name('option')
    statindex = 1
    for statname in statlt:

        if statname== stat:
            stats[statindex].click()
            #此处需要点击下其他位置才会生效选择的州名
            webQuteDriver.find_element_by_xpath('//*[@id="euPO"]').click()
            break
        statindex += 1
    #填充折扣
    webQuteDriver.find_element_by_xpath('//*[@id="spaRefNo"]').send_keys(shopInfo[0][3])
    #填充商品
    #shopInfo[0]:商品名、shopInfo[1]:数量、shopInfo[2]:单价、shopInfo[3]:折扣
    # 返回上层frame

    webQuteDriver.switch_to.parent_frame()

    '''
    for shop in shopInfo:
        WebDriverWait(webQuteDriver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="partNo"]')))
        webQuteDriver.find_element_by_xpath('//*[@id="partNo"]').clear()
        webQuteDriver.find_element_by_xpath('//*[@id="partNo"]').click()
        webQuteDriver.find_element_by_xpath('//*[@id="partNo"]').send_keys(shop[0])
        webQuteDriver.find_element_by_xpath('//*[@id="poQty"]').clear()
        webQuteDriver.find_element_by_xpath('//*[@id="poQty"]').click()
        webQuteDriver.find_element_by_xpath('//*[@id="poQty"]').send_keys(shop[1].replace(' ',''))
        webQuteDriver.find_element_by_xpath('//*[@id="btnGo"]/img').click()
        #等待商品填充完成，有问题得商品手动选择后填入得话会耗时很长所以为300s
        webQuteDriver.switch_to_frame('lineFrame')
        WebDriverWait(webQuteDriver, 300).until(EC.presence_of_element_located((By.XPATH, '//*[@id="partInfoSPan_{i}"]'.format(i = shopIndex))))
        WebDriverWait(webQuteDriver, 300).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="partInfoSPan_{i}"]'.format(i = shopIndex))))
        webQuteDriver.switch_to.parent_frame()
        time.sleep(1)
        '''
    #webQuteDriver.find_element_by_xpath('//*[@id="FlashID"]/object').click() 
    webQuteDriver.find_element_by_xpath('//*[@id="openBoxs"]/table/tbody/tr/td/img').click()
    quickSearchTable = ''
    for shop in shopInfo:
        quickSearchTable += shop[0] + ' ' + shop[1].replace(' ','') + '\r\n'
    webQuteDriver.find_element_by_xpath('//*[@id="quickSearchTable"]/tbody/tr/td/textarea').send_keys(quickSearchTable)
    webQuteDriver.find_element_by_xpath('//*[@id="quickSearchTable"]/tbody/tr/td/input[2]').click()

    webQuteDriver.switch_to.frame('lineFrame')
    WebDriverWait(webQuteDriver, 120).until(EC.presence_of_element_located((By.XPATH, '//*[@id="partInfoSPan_{i}"]'.format(i=len(shopInfo) - 1))))
    WebDriverWait(webQuteDriver, 120).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="partInfoSPan_{i}"]'.format(i=len(shopInfo) - 1))))
    shopIndex = 0
    for shop in shopInfo:
        webQuteDriver.find_element_by_xpath('//*[@id="lineForm[{i}].netPrice"]'.format(i=shopIndex)).send_keys(Keys.CONTROL,'a')
        webQuteDriver.find_element_by_xpath('//*[@id="lineForm[{i}].netPrice"]'.format(i=shopIndex)).send_keys(Keys.BACK_SPACE)
        time.sleep(0.1)
        webQuteDriver.find_element_by_xpath('//*[@id="lineForm[{i}].netPrice"]'.format(i = shopIndex)).send_keys(shop[2].replace(',','').split(' ')[0])
        print('Unit price ->:',shop[2].replace(',','').split(' ')[0])
        shopIndex += 1
    webQuteDriver.switch_to.parent_frame()
    webQuteDriver.find_element_by_xpath('//*[@id="batchValue"]').send_keys(order)
    webQuteDriver.find_element_by_xpath('//*[@id="change_batch"]').click()
def findWebQuteWind(driver):
    return  switchToWindwos('Web Quote',driver)
def findGfxWind(driver):
    return switchToWindwos('3dgfx Vendor Processing',driver)
def switchToGfxLoginWind(driver):
    while (1):
        driver = switchToWindwos('3Dgfx: Fulfillment',driver)
        if  driver.title == '3Dgfx: Fulfillment':
            break
        else:
            time.sleep(2)
    return driver
def serchOrder(driver,order):
    #手动点击
    #driver.find_element_by_class_name('tabs').click()
    driver.find_element_by_xpath('//*[@id="MAINFORM"]/table[4]/tbody/tr/td/table/tbody/tr/td[4]/table/tbody/tr/td[2]/span/a').click()

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'MAINFORM:plainmiddlebar:orderNum')))
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'MAINFORM:plainmiddlebar:orderNum')))

    driver.find_element_by_id('MAINFORM:plainmiddlebar:orderNum').clear()
    driver.find_element_by_id('MAINFORM:plainmiddlebar:orderNum').send_keys(order)
    driver.find_element_by_id('MAINFORM:plainmiddlebar:button').click()
def printAllWindow(driver):
    windows = driver.window_handles
    for SingleWindow in windows:
        driver.switch_to.window(SingleWindow)
        print("windows name is :",driver.title)
def getOrderInfo(driver):
    iD = []
    infoDir = {}
    shopInfoDir = []
    shopInfoList = []
    while(1):
        OrderIdElements = driver.find_elements_by_class_name('nostyle')
        if OrderIdElements != None:
            break
        time.sleep(1)
    for OrderIdElement in OrderIdElements:
        iD.append(OrderIdElement.text)
        #print(OrderIdElement.text)
    addrInfoHeadElements = driver.find_element_by_xpath('/html/body/form/table[7]/tbody/tr[1]')
    addrInfoElement = addrInfoHeadElements.find_elements_by_tag_name('td')

    addrElement = driver.find_element_by_class_name('vendor_detail_nostyle')
    addrInfoElements = addrElement.find_elements_by_tag_name('td')
    if 'Dell Customer Bill To' in addrInfoElement[0].text:
        addrInfoStr = addrInfoElements[1].text
        phoneAndEmail = addrInfoElements[2].text
    else:
        addrInfoStr = addrInfoElements[0].text
        phoneAndEmail = addrInfoElements[1].text
    global PRINTSWITCH
    if PRINTSWITCH == 1:
        print('addr info--> ', addrInfoStr)
    addrInfoList = addrInfoStr.split('\n')
    phoneAndEmailList = phoneAndEmail.split('\n')
    infoDir['custordr'] = iD[4]
    #print("len of addrInfoList->",len(addrInfoList))
    if len(addrInfoList) == 6:
        infoDir['company'] = addrInfoList[0]
        infoDir['user'] = addrInfoList[1]
        infoDir['addr'] = addrInfoList[2]
        infoDir['addr2'] = addrInfoList[3]
        infoDir['city'] = addrInfoList[4]
        infoDir['country'] = addrInfoList[5]
    elif len(addrInfoList) == 5:
        infoDir['company'] = addrInfoList[0]
        infoDir['user'] = addrInfoList[1]
        infoDir['addr'] = addrInfoList[2]
        #infoDir['addr2'] = addrInfoList[3]
        infoDir['city'] = addrInfoList[3]
        infoDir['country'] = addrInfoList[4]

    infoDir['phone'] = phoneAndEmailList[0]
    infoDir['email'] = phoneAndEmailList[1]

    #获取商品列表得头信息
    shopTableHeadElement = driver.find_element_by_xpath('//*[@id="MAINFORM"]/table[11]/thead/tr')
    shopTableHeadListElement = shopTableHeadElement.find_elements_by_tag_name('th')
    headIndex = 0
    shopNameIndex = 0
    shopQtyIndex = 0
    shopUnitCostIndex = 0
    shopLOReferenceNumberIndex = 0
    for shopTableHead in shopTableHeadListElement:
        if shopTableHead.text == 'Mfg PN':
            shopNameIndex = headIndex
        if shopTableHead.text == 'QTY':
            shopQtyIndex = headIndex
        if shopTableHead.text == 'Unit Cost':
            shopUnitCostIndex = headIndex
        if shopTableHead.text == 'LO Reference Number':
            shopLOReferenceNumberIndex = headIndex
        headIndex += 1
    shopTableElement = driver.find_element_by_xpath('/html/body/form/table[11]/tbody')
    shopElements = shopTableElement.find_elements_by_tag_name('tr')
    for shopElement in shopElements:
        shopList = shopElement.find_elements_by_tag_name('td')
        shopInfoDir.clear()
        shopInfoDir.append(shopList[shopNameIndex].text)
        shopInfoDir.append(shopList[shopQtyIndex].text)
        shopInfoDir.append(shopList[shopUnitCostIndex].text)
        shopInfoDir.append(shopList[shopLOReferenceNumberIndex].text)
        if PRINTSWITCH == 1:
            print(shopInfoDir)
        #列表使用值传递，要不然使用得全是最后一个列表得值
        shopInfoList.append(list(shopInfoDir))

    return infoDir,shopInfoList
def webQuetProc():
    chrome_options = Options()
    chrome_options.add_argument('--allow-running-insecure-content')
    webQuteDriver = webdriver.Ie()
    webQuteDriver.get("https://mycis.synnex.org/webQuote/viewMainPage.do")
    webQuoteLogin(webQuteDriver)
    return webQuteDriver
def gfxProc():
    gfxDriver = webdriver.Chrome()
    gfxDriver.get("https://gfx.dell.com/D3Dgfx/vendor/detail_update_success.faces")
    return gfxDriver
class myThread (threading.Thread):   #继承父类threading.Thread
    def __init__(self, gfxDriver, webQuteDriver,order):
        threading.Thread.__init__(self)
        self.gfxDriver = gfxDriver
        self.webQuteDriver = webQuteDriver
        self.order = order
    def run(self):                   #把要执行的代码写到run函数里面 线程在创建后会直接运行run函数
        Proc(self.gfxDriver,self.webQuteDriver,self.order)
def gfxLoginProc(gfxDriver):
    WebDriverWait(gfxDriver, 300).until(EC.presence_of_element_located((By.XPATH, '//*[@id="MAINFORM"]/center/table/tbody/tr[1]/td[2]/input')))
    WebDriverWait(gfxDriver, 300).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="MAINFORM"]/center/table/tbody/tr[1]/td[2]/input')))
    gfxDriver.find_element_by_xpath('//*[@id="MAINFORM"]/center/table/tbody/tr[1]/td[2]/input').clear()
    gfxDriver.find_element_by_xpath('//*[@id="MAINFORM"]/center/table/tbody/tr[1]/td[2]/input').send_keys('CAROLLIU')
    gfxDriver.find_element_by_xpath('//*[@id="MAINFORM"]/center/table/tbody/tr[2]/td[2]/input').send_keys('synnex-2020-2')
    gfxDriver.find_element_by_xpath('//*[@id="MAINFORM:button"]').click()
def webQuoteLogin(webQuteDriver):
    WebDriverWait(webQuteDriver, 300).until(EC.presence_of_element_located((By.XPATH,'//*[@id="wrap"]/div/div[2]/a')))
    WebDriverWait(webQuteDriver, 300).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="wrap"]/div/div[2]/a')))
    webQuteDriver.find_element_by_xpath('//*[@id="wrap"]/div/div[2]/a').click()

def Proc(gfxDriver,webQuteDriver,order):
    #order = input("start->")
    try:
        button['state'] = 'disabled'
        gfxDriver = findGfxWind(gfxDriver)
        webQuteDriver = findWebQuteWind(webQuteDriver)

        serchOrder(gfxDriver,order)
        addrInfoDir,shopInfo = getOrderInfo(gfxDriver)

        if PRINTSWITCH == 1:
            print(addrInfoDir,shopInfo,order)
        webQuteOperate(webQuteDriver,addrInfoDir,shopInfo,order)
        webQuteDriver.switch_to.default_content()
    except Exception as e:
        traceback.print_exc(file = open('log.txt','a+'))
        traceback.print_exc()
    button['state'] = 'normal'
def main(gfxDriver, webQuteDriver):
    global procThread
    order = orderEntry.get()
    procThread = myThread(gfxDriver,webQuteDriver,order)
    procThread.start()

if __name__ == '__main__':
    gfxDriver = gfxProc()
    #自动登录gfx网站
    gfxDriver = switchToGfxLoginWind(gfxDriver)
    gfxLoginProc(gfxDriver)

    webQuteDriver = webQuetProc()

    root = Tk()
    root.title("help")
    frame = Frame(root)
    frame.pack(padx=8, pady=8, ipadx=4)
    lab1 = Label(frame, text="NO:")
    lab1.grid(row=0, column=0, padx=5, pady=5, sticky=W)

    # 绑定对象到Entry

    orderVar = StringVar()
    orderEntry = Entry(frame, textvariable=orderVar)
    orderEntry.grid(row=0, column=1, sticky='ew', columnspan=2)


    button = Button(frame, text="录入", command=lambda: main(gfxDriver = gfxDriver,webQuteDriver = webQuteDriver), default='active')
    button.grid(row=2, column=1)

    button1 = Button(frame, text="停止", command=lambda: stopThread(tid=procThread._ident, cmd=SystemExit),
                    default='active')
    button1.grid(row=2, column=2)



    # 以下代码居中显示窗口

    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_reqwidth()) / 2
    y = (root.winfo_screenheight() - root.winfo_reqheight()) / 2
    root.geometry("+%d+%d" % (x, y))
    root.mainloop()

