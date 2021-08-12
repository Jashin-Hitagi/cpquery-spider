import os

import requests
from PIL import Image
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import time
import datetime
import xlsxwriter

# 打开浏览器
import getVeryCode

url = "http://cpquery.cnipa.gov.cn/"
excel_path = r'./excel'
if not os.path.exists(excel_path):  # 判断是否存在文件夹如果不存在则创建为文件夹
    os.makedirs(excel_path)
local_path = r'./file'
if not os.path.exists(local_path):  # 判断是否存在文件夹如果不存在则创建为文件夹
    os.makedirs(local_path)
localPath = r'./file/loadImg.png'
driver = webdriver.Firefox()
driver.get(url)


# 页面显示提示
def tips(text):
    print(text)
    driver.execute_script(
        "document.body.innerHTML=('<div style=\\'position:absolute;top:0px;left:0px;width:100%;text-align:center"
        ";background:#000;color:#fff;padding:8px;\\'>" + text + "</div>')+document.body.innerHTML")


# 截取验证码的图片并保存本地
def getAndSaveImg():
    element = driver.find_element_by_id('authImg')
    # 截取全屏图片
    driver.save_screenshot(r'./file/full.png')
    # 获取element的顶点坐标
    x_Piont = element.location['x']
    y_Piont = element.location['y']
    # 获取element的宽、高
    element_width = x_Piont + element.size['width']
    element_height = y_Piont + element.size['height']

    picture = Image.open(r'./file/full.png').convert("L")  # 转成灰度图

    picture = picture.crop((x_Piont, y_Piont, element_width, element_height))

    '''
    去掉截图下端的空白区域
    '''
    driver.execute_script(
        """
        $('#main').siblings().remove();
        $('#aside__wrapper').siblings().remove();
        $('.ui.sticky').siblings().remove();
        $('.follow-me').siblings().remove();
        $('img.ui.image').siblings().remove();
        """
    )
    out = picture.resize((70, 20))
    out.save(localPath)
    return out


# 验证码解析失败 刷新验证码并解析
def reLoadCode(newCode):
    newCode.click()  # 验证码解析失败 刷新验证码
    time.sleep(2)  # 等待验证码刷新
    newImg = getAndSaveImg()  # 获取刷新后的验证码
    newVeryCode = getVeryCode.getRealCode(newImg)  # 解析验证码
    return newVeryCode


# 判断是否有验证码输入错误弹框
def ifElementExist(elementName):
    try:
        driver.find_element_by_name(elementName)
        return True
    except:
        return False


# 解析验证码并跳转
def inputCode():
    WebDriverWait(driver, 600000).until(EC.presence_of_element_located((By.ID, 'authImg')))
    print('正在获取验证码')
    img = getAndSaveImg()
    veryCode = getVeryCode.getRealCode(img)
    getNewCode = driver.find_element_by_id('authImg')
    while veryCode == 404:  # 验证码识别失败
        veryCode = reLoadCode(getNewCode)  # 刷新验证码并解析
    queryUrl = "http://cpquery.cnipa.gov.cn/txnQueryOrdinaryPatents.do?select-key:sortcol=&select-key:sort=&select-key" \
               ":shenqingh=&select-key:zhuanlimc=&select-key:shenqingrxm=&select-key:zhuanlilx=&select-key:shenqingr_from" \
               "=&select-key:shenqingr_to=&select-key:dailirxm=&verycode=" + str(
        veryCode) + "&inner-flag:open-type=window&inner-flag:flowno" \
                    "=1616564232916 "
    driver.get(queryUrl)  # 在URL输入验证码结果 然后跳转 查询


# 等待登陆查询完成
WebDriverWait(driver, 600000).until(EC.presence_of_element_located((By.CLASS_NAME, 'login_butt')))
print('请手动登陆账号，在进入查询页面后系统自动开始导出')
# 解析验证码
inputCode()
time.sleep(3)
WebDriverWait(driver, 600000).until(EC.presence_of_element_located((By.CLASS_NAME, 'content_listx2')))  # 等待页面加载完毕

# 获取当前时间
now_time = datetime.datetime.now().strftime('%Y%m%d')
# 获取用户名
userName = "Test"

# 建立申请信息xlsx
xls_path_PatentApplyInfo = './excel/PatentApplyInfo.xlsx'
workbook_PatentApplyInfo = xlsxwriter.Workbook(xls_path_PatentApplyInfo)
booksheet_PatentApplyInfo = workbook_PatentApplyInfo.add_worksheet('data')

# 建立申请人xlsx
xls_path_PatentApplicantName = './excel/PatentApplicantName.xlsx'
workbook_PatentApplicantName = xlsxwriter.Workbook(xls_path_PatentApplicantName)
booksheet_PatentApplicantName = workbook_PatentApplicantName.add_worksheet('data')

# 建立优先权xlsx
xls_path_PatentPriority = './excel/PatentPriority.xlsx'
workbook_PatentPriority = xlsxwriter.Workbook(xls_path_PatentPriority)
booksheet_PatentPriority = workbook_PatentPriority.add_worksheet('data')

# 建立著录项目变更xlsx
xls_path_PatentItemRecordChange = './excel/PatentItemRecordChange.xlsx'
workbook_PatentItemRecordChange = xlsxwriter.Workbook(xls_path_PatentItemRecordChange)
booksheet_PatentItemRecordChange = workbook_PatentItemRecordChange.add_worksheet('data')

# 建立审查信息xlsx
xls_path_PatentCheckInfo = './excel/PatentCheckInfo.xlsx'
workbook_PatentCheckInfo = xlsxwriter.Workbook(xls_path_PatentCheckInfo)
booksheet_PatentCheckInfo = workbook_PatentCheckInfo.add_worksheet('data')

# 建立费用信息-应缴费信息xlsx
xls_path_PatentAmountCostInfo = './excel/PatentAmountCostInfo.xlsx'
workbook_PatentAmountCostInfo = xlsxwriter.Workbook(xls_path_PatentAmountCostInfo)
booksheet_PatentAmountCostInfo = workbook_PatentAmountCostInfo.add_worksheet('data')

# 建立费用信息-应缴费信息xlsx
xls_path_ChildPatentAmountCostInfo = './excel/ChildPatentAmountCostInfo.xlsx'
workbook_ChildPatentAmountCostInfo = xlsxwriter.Workbook(xls_path_ChildPatentAmountCostInfo)
booksheet_ChildPatentAmountCostInfo = workbook_ChildPatentAmountCostInfo.add_worksheet('data')

# 建立费用信息-已缴费信息xlsx
xls_path_PatentPaidInfo = './excel/PatentPaidInfo.xlsx'
workbook_PatentPaidInfo = xlsxwriter.Workbook(xls_path_PatentPaidInfo)
booksheet_PatentPaidInfo = workbook_PatentPaidInfo.add_worksheet('data')

# 建立费用信息-已缴费信息xlsx
xls_path_ChildPatentPaidInfo = './excel/ChildPatentPaidInfo.xlsx'
workbook_ChildPatentPaidInfo = xlsxwriter.Workbook(xls_path_ChildPatentPaidInfo)
booksheet_ChildPatentPaidInfo = workbook_ChildPatentPaidInfo.add_worksheet('data')

# 建立费用信息-冲红信息xlsx
xls_path_PatentRedInfo = './excel/PatentRedInfo.xlsx'
workbook_PatentRedInfo = xlsxwriter.Workbook(xls_path_PatentRedInfo)
booksheet_PatentRedInfo = workbook_PatentRedInfo.add_worksheet('data')

# 建立费用信息-冲红信息xlsx
xls_path_ChildPatentRedInfo = './excel/ChildPatentRedInfo.xlsx'
workbook_ChildPatentRedInfo = xlsxwriter.Workbook(xls_path_ChildPatentRedInfo)
booksheet_ChildPatentRedInfo = workbook_ChildPatentRedInfo.add_worksheet('data')

# 建立费用信息-退费信息xlsx
xls_path_PatentRefundInfo = './excel/PatentRefundInfo.xlsx'
workbook_PatentRefundInfo = xlsxwriter.Workbook(xls_path_PatentRefundInfo)
booksheet_PatentRefundInfo = workbook_PatentRefundInfo.add_worksheet('data')

# 建立费用信息-退费信息xlsx
xls_path_ChildPatentRefundInfo = './excel/ChildPatentRefundInfo.xlsx'
workbook_ChildPatentRefundInfo = xlsxwriter.Workbook(xls_path_ChildPatentRefundInfo)
booksheet_ChildPatentRefundInfo = workbook_ChildPatentRefundInfo.add_worksheet('data')

# 建立费用信息-滞纳金信息xlsx
xls_path_PatentLateFeeInfo = './excel/PatentLateFeeInfo.xlsx'
workbook_PatentLateFeeInfo = xlsxwriter.Workbook(xls_path_PatentLateFeeInfo)
booksheet_PatentLateFeeInfo = workbook_PatentLateFeeInfo.add_worksheet('data')

# 建立费用信息-滞纳金信息xlsx
xls_path_ChildPatentLateFeeInfo = './excel/ChildPatentLateFeeInfo.xlsx'
workbook_ChildPatentLateFeeInfo = xlsxwriter.Workbook(xls_path_ChildPatentLateFeeInfo)
booksheet_ChildPatentLateFeeInfo = workbook_ChildPatentLateFeeInfo.add_worksheet('data')

# 建立费用信息-收据发文信息xlsx
xls_path_PatentReceiptPostInfo = './excel/PatentReceiptPostInfo.xlsx'
workbook_PatentReceiptPostInfo = xlsxwriter.Workbook(xls_path_PatentReceiptPostInfo)
booksheet_PatentReceiptPostInfo = workbook_PatentReceiptPostInfo.add_worksheet('data')

# 建立费用信息-收据发文信息-收据发文信息xlsx
xls_path_ChildPatentReceiptPostInfo = './excel/ChildPatentReceiptPostInfo.xlsx'
workbook_ChildPatentReceiptPostInfo = xlsxwriter.Workbook(xls_path_ChildPatentReceiptPostInfo)
booksheet_ChildPatentReceiptPostInfo = workbook_ChildPatentReceiptPostInfo.add_worksheet('data')

# 建立费用信息-收据发文信息-缴费信息xlsx
xls_path_ChildPatentReceiptCostInfo = './excel/child_patent_receipt_post_info.xlsx'
workbook_ChildPatentReceiptCostInfo = xlsxwriter.Workbook(xls_path_ChildPatentReceiptCostInfo)
booksheet_ChildPatentReceiptCostInfo = workbook_ChildPatentReceiptCostInfo.add_worksheet('data')

# 建立发文信息-通知书发文xlsx
xls_path_PatentNoticeIssued = './excel/PatentNoticeIssued.xlsx'
workbook_PatentNoticeIssued = xlsxwriter.Workbook(xls_path_PatentNoticeIssued)
booksheet_PatentNoticeIssued = workbook_PatentNoticeIssued.add_worksheet('data')

# 建立发文信息-通知书发文-通知书发文xlsx
xls_path_ChildPatentNoticeIssued = './excel/ChildPatentNoticeIssued.xlsx'
workbook_ChildPatentNoticeIssued = xlsxwriter.Workbook(xls_path_ChildPatentNoticeIssued)
booksheet_ChildPatentNoticeIssued = workbook_ChildPatentNoticeIssued.add_worksheet('data')

# 建立发文信息-专利证书xlsx
xls_path_PatentCertificate = './excel/PatentCertificate.xlsx'
workbook_PatentCertificate = xlsxwriter.Workbook(xls_path_PatentCertificate)
booksheet_PatentCertificate = workbook_PatentCertificate.add_worksheet('data')

# 建立发文信息-专利证书-专利证书xlsx
xls_path_ChildPatentCertificate = './excel/ChildPatentCertificate.xlsx'
workbook_ChildPatentCertificate = xlsxwriter.Workbook(xls_path_ChildPatentCertificate)
booksheet_ChildPatentCertificate = workbook_ChildPatentCertificate.add_worksheet('data')

# 建立发文信息-退信xlsx
xls_path_PatentReturnLetter = './excel/PatentReturnLetter.xlsx'
workbook_PatentReturnLetter = xlsxwriter.Workbook(xls_path_PatentReturnLetter)
booksheet_PatentReturnLetter = workbook_PatentReturnLetter.add_worksheet('data')

# 建立发文信息-退信-退信xlsx
xls_path_ChildPatentReturnLetter = './excel/ChildPatentReturnLetter.xlsx'
workbook_ChildPatentReturnLetter = xlsxwriter.Workbook(xls_path_ChildPatentReturnLetter)
booksheet_ChildPatentReturnLetter = workbook_ChildPatentReturnLetter.add_worksheet('data')

# 建立公布公告-事务公告xlsx
xls_path_PatentAffairsReport = './excel/PatentAffairsReport.xlsx'
workbook_PatentAffairsReport = xlsxwriter.Workbook(xls_path_PatentAffairsReport)
booksheet_PatentAffairsReport = workbook_PatentAffairsReport.add_worksheet('data')

# 建立公布公告-事务公告-事务公告xlsx
xls_path_ChildPatentAffairsReport = './excel/ChildPatentAffairsReport.xlsx'
workbook_ChildPatentAffairsReport = xlsxwriter.Workbook(xls_path_ChildPatentAffairsReport)
booksheet_ChildPatentAffairsReport = workbook_ChildPatentAffairsReport.add_worksheet('data')

# 建立公布公告-发明公布/授权公告xlsx
xls_path_PatentInventionReport = './excel/PatentInventionReport.xlsx'
workbook_PatentInventionReport = xlsxwriter.Workbook(xls_path_PatentInventionReport)
booksheet_PatentInventionReport = workbook_PatentInventionReport.add_worksheet('data')

# 建立公布公告-发明公布/授权公告-发明公布/授权公告xlsx
xls_path_ChildPatentInventionReport = './excel/ChildPatentInventionReport.xlsx'
workbook_ChildPatentInventionReport = xlsxwriter.Workbook(xls_path_ChildPatentInventionReport)
booksheet_ChildPatentInventionReport = workbook_ChildPatentInventionReport.add_worksheet('data')

# 建立公布公告-同族案件信息xlsx
xls_path_PatentFamilyCaseInfo = './excel/PatentFamilyCaseInfo.xlsx'
workbook_PatentFamilyCaseInfo = xlsxwriter.Workbook(xls_path_PatentFamilyCaseInfo)
booksheet_PatentFamilyCaseInfo = workbook_PatentFamilyCaseInfo.add_worksheet('data')

# 建立公布公告-同族案件信息-同族案件xlsx
xls_path_PatentFamilyCaseDetail = './excel/PatentFamilyCaseDetail.xlsx'
workbook_PatentFamilyCaseDetail = xlsxwriter.Workbook(xls_path_PatentFamilyCaseDetail)
booksheet_PatentFamilyCaseDetail = workbook_PatentFamilyCaseDetail.add_worksheet('data')

# 建立专利案件xlsx
xls_path_PatentCaseInfo = './excel/PatentCaseInfo.xlsx'
workbook_PatentCaseInfo = xlsxwriter.Workbook(xls_path_PatentCaseInfo)
booksheet_PatentCaseInfo = workbook_PatentCaseInfo.add_worksheet('data')

shenqingNo = 1  # 专利顺序编号
startRow = 1  # 每页的第一条专利顺序编号
PatentApplicantNameRow = 1  # 申请人excel的行编号
PatentPriorityRow = 1  # 优先权excel的行编号
PatentItemRecordChangeRow = 1  # 著录项目变更excel的行编号
fyChildSheetRowList = [1] * 7  # 费用信息的行编号
fwChildSheetRowList = [1] * 7  # 发文信息的行编号
gbChildSheetRowList = [1] * 7  # 公布公告的行编号
PatentFamilyCaseDetailRow = 1  # 同族案件信息的行编号

time.sleep(5)  # 等待页面加载
# 获取总页面数
total_page = int(
    driver.find_elements_by_css_selector('.form-control')[0].find_element_by_xpath('..').text.replace('/',
                                                                                                      '').strip())
print('总页数为：', total_page)

# 打开新标签抓取详细信息
driver.execute_script('''window.open("about:blank","_blank");''')
driver.switch_to.window(driver.window_handles[0])

# 遍历每一页
for i in range(total_page):
    if i != 0:
        driver.find_element_by_class_name('pagination').find_elements_by_tag_name('li')[3].click()  # 点击下一页
    print("正在导出，当前在第" + str(i+1) + "页")
    WebDriverWait(driver, 600000).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'content_listx_patent')))  # 等待页面加载完毕
    time.sleep(10)  # 等待页面加载
    now_page = int(
        driver.find_elements_by_css_selector('.form-control')[0].get_attribute('value'))
    PatentCaseInfoTr = \
        driver.find_elements_by_css_selector('.select_box')[0].find_element_by_tag_name(
            'table').find_elements_by_tag_name(
            'tr')[0]
    tdList = PatentCaseInfoTr.find_elements_by_tag_name('td')  # 取得title行
    # 设置title行
    for tr in range(len(tdList)):
        booksheet_PatentCaseInfo.write(0, tr, tdList[tr].text)
        if tr == 2:
            booksheet_PatentCaseInfo.write(0, tr, "姓名或名称")
        if tr == len(tdList) - 1:
            booksheet_PatentCaseInfo.write(0, tr + 1, "主分类号")
            booksheet_PatentCaseInfo.write(0, tr + 2, "分案提交日")
    es = driver.find_elements_by_css_selector('.content_listx2')
    c = len(es)
    shenqinghIDList = []
    shenqinghNameList = []
    shenqinghTypeList = []
    # 遍历每一行数据
    for j in range(c):
        # 导出第J行数据
        item = driver.find_elements_by_css_selector('.content_listx2')[j].find_element_by_class_name(
            'content_listx_patent').find_elements_by_css_selector('td')
        for k in range(len(item)):
            booksheet_PatentCaseInfo.write(i*20+j+1, k, item[k].text)
        shenqinghID = driver.find_elements_by_css_selector('.content_listx2')[j].find_element_by_class_name(
            'content_listx_patent').find_elements_by_css_selector('td')[0].text
        shenqinghIDList.append(shenqinghID)
        shenqinghName = driver.find_elements_by_css_selector('.content_listx2')[j].find_element_by_class_name(
            'content_listx_patent').find_elements_by_css_selector('td')[1].text
        shenqinghNameList.append(shenqinghName)
        shenqinghType = driver.find_elements_by_css_selector('.content_listx2')[j].find_element_by_class_name(
            'content_listx_patent').find_elements_by_css_selector('td')[5].text
        shenqinghTypeList.append(shenqinghType)
    driver.switch_to.window(driver.window_handles[1])  # 切换标签
    for j in range(c):
        print("正在导出，当前在第" + str(i + 1) + "页,第" + str(j + 1) + "行,总顺序第" + str(shenqingNo) + "条")

        # 进入专利号-专利详情页面
        shenqinghID = shenqinghIDList[j]
        shenqinghName = shenqinghNameList[j]
        shenqinghType = shenqinghTypeList[j]

        driver.get('http://cpquery.cnipa.gov.cn/txnQueryBibliographicData.do?select-key:shenqingh=' + shenqinghID)
        WebDriverWait(driver, 600000).until(EC.presence_of_element_located((By.CLASS_NAME, 'tab_list')))  # 等待页面加载完毕
        time.sleep(5)  # 等待页面加载完毕

        # 展开下拉表格并写入标题
        def pullTable(element, targetSheet):
            try:
                element.find_elements_by_tag_name('h2')[0].find_element_by_class_name('draw_down').click()
                table = element.find_element_by_css_selector('.imfor_table_grid')
                trList = table.find_elements_by_tag_name('tr')
                thList = trList[0].find_elements_by_tag_name('th')  # 标题行
                lineCount = len(thList)  # 该数据占的列数
                for p in range(lineCount):
                    targetSheet.write(0, p, thList[p].text)  # 写入EXCEL标题
                targetSheet.write(0, lineCount, "申请号/专利号")
            except NoSuchElementException as meg:
                table = element.find_element_by_css_selector('.imfor_table_grid')
                trList = table.find_elements_by_tag_name('tr')
                thList = trList[0].find_elements_by_tag_name('th')  # 标题行
                lineCount = len(thList)  # 该数据占的列数
                for p in range(lineCount):
                    targetSheet.write(0, p, thList[p].text)  # 写入EXCEL标题
                targetSheet.write(0, lineCount, "申请号/专利号")
                # print('找不到按钮:', meg)

        # 展开下拉列表并写入标题
        def pullList(element, colNo, targetSheet):
            try:
                element.find_elements_by_tag_name('h2')[0].find_element_by_class_name('draw_down').click()
                titleList = element.find_elements_by_css_selector('.td1')
                for t in range(len(titleList)):
                    title = str(titleList[t].text).strip('：').strip()
                    targetSheet.write(0, colNo + t, title)  # 写入EXCEL标题
            except NoSuchElementException as meg:
                # print('找不到按钮:', meg)
                titleList = element.find_elements_by_css_selector('.td1')
                for t in range(len(titleList)):
                    title = str(titleList[t].text).strip('：').strip()
                    targetSheet.write(0, colNo + t, title)  # 写入EXCEL标题

        # 写入费用信息 发文信息 公布公告
        def writeFyFwChild(element, targetSheet, rowNo, getTabId):
            global fyChildSheetRowList
            global fwChildSheetRowList
            global gbChildSheetRowList
            table = element.find_element_by_css_selector('.imfor_table_grid')
            trList = table.find_elements_by_tag_name('tr')
            thList = trList[0].find_elements_by_tag_name('th')  # 标题行
            lineCount = len(thList)  # 该数据占的列宽
            for p in range(1, len(trList)):
                tdList = trList[p].find_elements_by_tag_name('td')
                for k in range(lineCount):
                    if thList[0].text == "" and k == 0:
                        resultData = ""
                    else:
                        resultData = tdList[k].find_elements_by_tag_name('span')[0] \
                            .get_attribute('title')
                    if getTabId == 'fyxx':
                        targetSheet.write(fyChildSheetRowList[rowNo], k, resultData)  # 写入EXCEL内容
                    if getTabId == 'fwxx':
                        targetSheet.write(fwChildSheetRowList[rowNo], k, resultData)  # 写入EXCEL内容
                    if getTabId == 'gbgg':
                        targetSheet.write(gbChildSheetRowList[rowNo], k, resultData)  # 写入EXCEL内容
                    # print(resultData)
                # 写入EXCEL主键
                if getTabId == 'fyxx':
                    targetSheet.write(fyChildSheetRowList[rowNo], lineCount, shenqinghID)
                    fyChildSheetRowList[rowNo] += 1
                if getTabId == 'fwxx':
                    targetSheet.write(fwChildSheetRowList[rowNo], lineCount, shenqinghID)
                    fwChildSheetRowList[rowNo] += 1
                if getTabId == 'gbgg':
                    targetSheet.write(gbChildSheetRowList[rowNo], lineCount, shenqinghID)
                    gbChildSheetRowList[rowNo] += 1

        # 根据tabId导出数据
        def outputData(inputId):
            # 遍历申请信息tab的数据
            if inputId == 'jbxx':
                # 添加审查信息
                booksheet_PatentCheckInfo.write(0, 0, "申请号/专利号")
                booksheet_PatentCheckInfo.write(shenqingNo, 0, shenqinghID)  # 写入专利ID
                booksheet_PatentCheckInfo.write(0, 1, "发明名称")
                booksheet_PatentCheckInfo.write(shenqingNo, 1, shenqinghName)  # 写入专利名称
                lineNo = 0  # 申请信息EXCEL列编号
                global PatentApplicantNameRow
                global PatentPriorityRow
                global PatentItemRecordChangeRow
                # tip('申请信息:')
                allIm = driver.find_elements_by_css_selector('.imfor_part1')
                # tip('著录项目信息---------')
                tableList = allIm[0].find_elements_by_css_selector('.imfor_table_grid')
                for k in range(len(tableList)):
                    trList = tableList[k].find_elements_by_tag_name('tr')
                    for tr in trList:
                        title = str(tr.find_elements_by_tag_name('td')[0].text).strip('：').strip()
                        booksheet_PatentApplyInfo.write(0, lineNo, title)
                        resultData = tr.find_elements_by_tag_name('td')[1].text
                        booksheet_PatentApplyInfo.write(shenqingNo, lineNo, resultData)  # 写入EXCEL内容
                        if title == '主分类号':
                            booksheet_PatentCaseInfo.write(shenqingNo, 8, resultData)
                        if title == '分案提交日':
                            booksheet_PatentCaseInfo.write(shenqingNo, 9, resultData)
                        lineNo += 1  # 换下一列
                        # print(resultData)
                # tip('申请人---------')
                pullTable(allIm[1], booksheet_PatentApplicantName)
                table = allIm[1].find_element_by_css_selector('.imfor_table_grid')
                trList = table.find_elements_by_tag_name('tr')
                tdList = trList[0].find_elements_by_tag_name('th')  # 标题行
                lineCount = len(tdList)  # 该数据占的列数
                for p in range(1, len(trList)):
                    tdList = trList[p].find_elements_by_tag_name('td')
                    for k in range(lineCount):
                        resultData = tdList[k].find_element_by_tag_name('span') \
                            .get_attribute('title')
                        booksheet_PatentApplicantName.write(PatentApplicantNameRow, k, resultData)  # 写入EXCEL内容
                    # 写入EXCEL主键
                    booksheet_PatentApplicantName.write(PatentApplicantNameRow, lineCount, shenqinghID)
                    PatentApplicantNameRow += 1
                    # print(resultData)
                # tip('发明人/设计人---------')
                table = allIm[2].find_element_by_css_selector('.imfor_table_grid')
                trList = table.find_elements_by_tag_name('tr')
                for tr in trList:
                    title = str(tr.find_elements_by_tag_name('td')[0].text).strip('：').strip()
                    booksheet_PatentApplyInfo.write(0, lineNo, title)  # 写入EXCEL标题
                    resultData = tr.find_elements_by_tag_name('td')[1].text
                    booksheet_PatentApplyInfo.write(shenqingNo, lineNo, resultData)  # 写入EXCEL
                    lineNo += 1  # 换下一列
                    # print(resultData)
                # tip('联系人---------')
                tableList = allIm[3].find_elements_by_css_selector('.imfor_table_grid')
                pullList(allIm[3], lineNo, booksheet_PatentApplyInfo)
                for k in range(len(tableList)):
                    trList = tableList[k].find_elements_by_tag_name('tr')
                    for tr in trList:
                        if lineNo == 7:
                            booksheet_PatentApplyInfo.write(0, 7, "联系人姓名")  # 修改EXCEL标题
                        resultData = tr.find_elements_by_tag_name('td')[1].text
                        booksheet_PatentApplyInfo.write(shenqingNo, lineNo, resultData)  # 写入EXCEL内容
                        lineNo += 1  # 换下一列
                        # print(resultData)
                # tip('代理情况---------')
                tableList = allIm[4].find_elements_by_css_selector('.imfor_table_grid')
                for k in range(len(tableList)):
                    trList = tableList[k].find_elements_by_tag_name('tr')
                    for tr in trList:
                        title = str(tr.find_elements_by_tag_name('td')[0].text).strip('：').strip()
                        booksheet_PatentApplyInfo.write(0, lineNo, title)  # 写入EXCEL标题
                        resultData = tr.find_elements_by_tag_name('td')[1].text
                        booksheet_PatentApplyInfo.write(shenqingNo, lineNo, resultData)  # 写入EXCEL内容
                        lineNo += 1  # 换下一列
                        # print(resultData)
                # tip('优先权---------')
                pullTable(allIm[5], booksheet_PatentPriority)
                table = allIm[5].find_element_by_css_selector('.imfor_table_grid')
                trList = table.find_elements_by_tag_name('tr')
                tdList = trList[0].find_elements_by_tag_name('th')  # 标题行
                lineCount = len(tdList)  # 该数据占的列数
                for p in range(1, len(trList)):
                    tdList = trList[p].find_elements_by_tag_name('td')
                    for k in range(lineCount):
                        resultData = tdList[k].find_element_by_tag_name('span') \
                            .get_attribute('title')
                        booksheet_PatentPriority.write(PatentPriorityRow, k, resultData)  # 写入EXCEL内容
                    # 写入EXCEL主键
                    booksheet_PatentPriority.write(PatentApplicantNameRow, lineCount, shenqinghID)
                    PatentPriorityRow += 1
                    # print(resultData)
                # tip('申请国际阶段---------')
                tableList = allIm[6].find_elements_by_css_selector('.imfor_table_grid')
                pullList(allIm[6], lineNo, booksheet_PatentApplyInfo)
                for k in range(len(tableList)):
                    trList = tableList[k].find_elements_by_tag_name('tr')
                    for tr in trList:
                        resultData = tr.find_elements_by_tag_name('td')[1].text
                        booksheet_PatentApplyInfo.write(shenqingNo, lineNo, resultData)  # 写入EXCEL内容
                        lineNo += 1  # 换下一列
                        # print(resultData)
                # 添加专利类型列
                booksheet_PatentApplyInfo.write(0, lineNo, "专利类型")  # 写入EXCEL内容
                booksheet_PatentApplyInfo.write(shenqingNo, lineNo, shenqinghType)  # 写入EXCEL内容
                # tip('著录项目变更---------')
                pullTable(allIm[7], booksheet_PatentItemRecordChange)
                table = allIm[7].find_element_by_css_selector('.imfor_table_grid')
                trList = table.find_elements_by_tag_name('tr')
                tdList = trList[0].find_elements_by_tag_name('th')  # 标题行
                lineCount = len(tdList)  # 该数据占的列数
                for p in range(1, len(trList)):
                    tdList = trList[p].find_elements_by_tag_name('td')
                    for k in range(lineCount):
                        resultData = tdList[k].find_element_by_tag_name('span') \
                            .get_attribute('title')
                        booksheet_PatentItemRecordChange.write(PatentItemRecordChangeRow, k, resultData)  # 写入EXCEL内容
                    # 写入EXCEL主键
                    booksheet_PatentItemRecordChange.write(PatentItemRecordChangeRow, lineCount, shenqinghID)
                    PatentItemRecordChangeRow += 1
                    # print(resultData)
                # print('当前列编号', lineNo)

            # 遍历费用信息和发文信息tab的数据
            if inputId == 'fyxx' or inputId == 'fwxx':
                fySheetList = [booksheet_PatentAmountCostInfo, booksheet_PatentPaidInfo, booksheet_PatentRedInfo,
                               booksheet_PatentRefundInfo, booksheet_PatentLateFeeInfo, booksheet_PatentReceiptPostInfo]
                fyChildSheetList = [booksheet_ChildPatentAmountCostInfo, booksheet_ChildPatentPaidInfo,
                                    booksheet_ChildPatentRedInfo, booksheet_ChildPatentRefundInfo,
                                    booksheet_ChildPatentLateFeeInfo, booksheet_ChildPatentReceiptPostInfo,
                                    booksheet_ChildPatentReceiptCostInfo]
                fwSheetList = [booksheet_PatentNoticeIssued, booksheet_PatentCertificate, booksheet_PatentReturnLetter]
                fwChildSheetList = [booksheet_ChildPatentNoticeIssued,
                                    booksheet_ChildPatentCertificate, booksheet_ChildPatentReturnLetter]
                # tip('费用信息和发文信息:')
                allIm = driver.find_elements_by_css_selector('.imfor_part1')
                for k in range(len(allIm)):
                    if inputId == 'fyxx':
                        pullTable(allIm[k], fyChildSheetList[k])
                        writeFyFwChild(allIm[k], fyChildSheetList[k], k, inputId)
                        if k == 5:
                            pullTable(allIm[k], fyChildSheetList[k + 1])
                            writeFyFwChild(allIm[k], fyChildSheetList[k + 1], k + 1, inputId)
                        fySheetList[k].write(0, 0, "申请号/专利号")
                        fySheetList[k].write(shenqingNo, 0, shenqinghID)  # 写入专利ID
                        fySheetList[k].write(0, 1, "发明名称")
                        fySheetList[k].write(shenqingNo, 1, shenqinghName)  # 写入专利名称
                    if inputId == 'fwxx':
                        pullTable(allIm[k], fwChildSheetList[k])
                        writeFyFwChild(allIm[k], fwChildSheetList[k], k, inputId)
                        fwSheetList[k].write(0, 0, "申请号/专利号")
                        fwSheetList[k].write(shenqingNo, 0, shenqinghID)  # 写入专利ID
                        fwSheetList[k].write(0, 1, "发明名称")
                        fwSheetList[k].write(shenqingNo, 1, shenqinghName)  # 写入专利名称

            # 遍历公布公告tab的数据
            if inputId == 'gbgg':
                gbSheetList = [booksheet_PatentInventionReport, booksheet_PatentAffairsReport]
                gbChildSheetList = [booksheet_ChildPatentInventionReport, booksheet_ChildPatentAffairsReport]
                allIm = driver.find_elements_by_css_selector('.imfor_part1')
                for k in range(len(allIm)):
                    pullTable(allIm[k], gbChildSheetList[k])
                    writeFyFwChild(allIm[k], gbChildSheetList[k], k, inputId)
                    gbSheetList[k].write(0, 0, "申请号/专利号")
                    gbSheetList[k].write(shenqingNo, 0, shenqinghID)  # 写入专利ID
                    gbSheetList[k].write(0, 1, "发明名称")
                    gbSheetList[k].write(shenqingNo, 1, shenqinghName)  # 写入专利名称

            # 遍历同族案件信息tab的数据
            if inputId == 'tzzlxx':
                global PatentFamilyCaseDetailRow
                # 写入标题
                booksheet_PatentFamilyCaseDetail.write(0, 0, "申请号/专利号")
                booksheet_PatentFamilyCaseDetail.write(0, 1, "同族案件申请号")
                booksheet_PatentFamilyCaseDetail.write(0, 2, "公开号")
                booksheet_PatentFamilyCaseDetail.write(0, 3, "公开日")
                booksheet_PatentFamilyCaseDetail.write(0, 4, "申请日")
                booksheet_PatentFamilyCaseInfo.write(0, 0, "申请号/专利号")
                booksheet_PatentFamilyCaseInfo.write(shenqingNo, 0, shenqinghID)  # 写入专利ID
                booksheet_PatentFamilyCaseInfo.write(0, 1, "发明名称")
                booksheet_PatentFamilyCaseInfo.write(shenqingNo, 1, shenqinghName)  # 写入专利名称
                WebDriverWait(driver, 600000).until(EC.presence_of_element_located((By.CLASS_NAME, 'tab_list')))
                # tip('同族案件信息:')
                try:
                    table = driver.find_element_by_id('tableContent')
                    trList = table.find_elements_by_tag_name('tr')
                    for k in range(len(trList)):
                        # EXCEL列编号
                        lineNo = 1
                        tdList = trList[k].find_elements_by_tag_name('td')
                        for td in tdList:
                            li1 = td.find_elements_by_tag_name('li')[0]
                            spanList = li1.find_elements_by_tag_name('span')
                            string_list = []
                            for span in spanList:
                                string_list.append(span.text)
                            resultData = ' '.join(string_list)
                            booksheet_PatentFamilyCaseDetail.write(PatentFamilyCaseDetailRow, lineNo, resultData)
                            lineNo += 1
                            # print('申请号：', resultData)
                            li2 = td.find_elements_by_tag_name('li')[1]
                            pList = li2.find_elements_by_tag_name('p')
                            for p in range(len(pList)):
                                spanList = pList[p].find_elements_by_tag_name('span')
                                string_list = []
                                for span in spanList:
                                    string_list.append(span.text)
                                resultData = ' '.join(string_list)
                                booksheet_PatentFamilyCaseDetail.write(PatentFamilyCaseDetailRow, lineNo + p,
                                                                       resultData)
                                # print('公开：', resultData)
                        booksheet_PatentFamilyCaseDetail.write(PatentFamilyCaseDetailRow, 0, shenqinghID)
                        PatentFamilyCaseDetailRow += 1
                except NoSuchElementException as meg:
                    print('无同族案件:', meg)


        # 获取详细信息的tab个数
        tab = driver.find_element_by_xpath('//*[@class="tab_list"]/ul')
        tabs = tab.find_elements_by_css_selector('.tab_top')
        tabCount = len(tabs) + 1  # tab个数
        tabNo = 1  # tab顺序
        # 循环tab导出数据
        for qu in range(2, tabCount + 1):
            WebDriverWait(driver, 600000).until(EC.presence_of_element_located((By.CLASS_NAME, 'tab_box')))
            # 获取本tab的ID
            thisTab = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[1]/ul/li[' + str(tabNo) + ']')
            tabID = thisTab.get_attribute('id')
            tabNo = qu + 1
            # 遍历申请信息tab的数据
            outputData(tabID)

            # 如果当前tab不是最后一个 跳转到下一个tab标签
            if tabNo <= tabCount:
                driver.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div[1]/ul/li[' + str(qu + 1) + ']').click()  # 跳转到下一个tab标签
                # 等待页面加载完毕
                if qu == 2:
                    WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, 'sjfw')))
                    time.sleep(4)
                elif qu == 3:
                    WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, 'txid')))
                    time.sleep(4)
                elif qu == 4:
                    WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.ID, 'swggid')))
                    time.sleep(4)
                elif qu == 5:
                    time.sleep(5)
            # 如果当前tab为最后一个 返回至查询页
            else:
                shenqingNo += 1
                break
    driver.switch_to.window(driver.window_handles[0])
workbook_PatentApplyInfo.close()
workbook_PatentPriority.close()
workbook_PatentCaseInfo.close()
workbook_PatentApplicantName.close()
workbook_PatentFamilyCaseDetail.close()
workbook_PatentFamilyCaseInfo.close()
workbook_PatentCheckInfo.close()
workbook_PatentItemRecordChange.close()
workbook_ChildPatentInventionReport.close()
workbook_PatentInventionReport.close()
workbook_ChildPatentAffairsReport.close()
workbook_PatentAffairsReport.close()
workbook_ChildPatentReturnLetter.close()
workbook_PatentReturnLetter.close()
workbook_ChildPatentCertificate.close()
workbook_PatentCertificate.close()
workbook_ChildPatentNoticeIssued.close()
workbook_PatentNoticeIssued.close()
workbook_ChildPatentReceiptCostInfo.close()
workbook_ChildPatentReceiptPostInfo.close()
workbook_PatentReceiptPostInfo.close()
workbook_ChildPatentLateFeeInfo.close()
workbook_PatentLateFeeInfo.close()
workbook_ChildPatentRefundInfo.close()
workbook_PatentRefundInfo.close()
workbook_ChildPatentRedInfo.close()
workbook_PatentRedInfo.close()
workbook_ChildPatentPaidInfo.close()
workbook_PatentPaidInfo.close()
workbook_ChildPatentAmountCostInfo.close()
workbook_PatentAmountCostInfo.close()

local_path = r'./excel'
url = ''

for root, dirs, files in os.walk(local_path, topdown=True):
    for file in files:
        filePath = os.path.join(root, file)
        sendFile = {"file": open(filePath, "rb")}
        r = requests.post(url, files=sendFile)
        print(r.text)

tips('数据导出结束!')
time.sleep(5)
driver.close()
