# -*- coding: utf-8 -*-
import time
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
# 以下两项被注释的原因可详见具体代码中的注释。
# from webdriver_manager.microsoft import IEDriverManager
# from webdriver_manager.opera import OperaDriverManager
import xlwt
import uuid
from bs4 import BeautifulSoup
import traceback
import re

class xls_generate:
    def __init__(self) -> None:
        pass

    def centerStyle(self, bold=False, height=200):
        style = xlwt.XFStyle()

        font = xlwt.Font()
        font.name = "微软雅黑"
        font.bold = bold
        font.underline = False
        font.italic = False
        font.colour_index = 0
        font.height = height
        style.font = font

        align = xlwt.Alignment()
        align.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
        align.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
        style.alignment = align

        # 边框
        border = xlwt.Borders()  # 给单元格加框线
        border.left = xlwt.Borders.THIN  # 左
        border.top = xlwt.Borders.THIN  # 上
        border.right = xlwt.Borders.THIN  # 右
        border.bottom = xlwt.Borders.THIN  # 下
        border.left_colour = 0x40  # 边框线颜色
        border.right_colour = 0x40
        border.top_colour = 0x40
        border.bottom_colour = 0x40
        style.borders = border

        return style

    def get_style(self, bold=False):
        style = xlwt.XFStyle()

        font = xlwt.Font()
        font.name = "微软雅黑"
        font.bold = bold
        font.underline = False
        font.italic = False
        font.colour_index = 0
        font.height = 200  # 200为10号字体
        style.font = font

        # 边框
        border = xlwt.Borders()  # 给单元格加框线
        border.left = xlwt.Borders.THIN  # 左
        border.top = xlwt.Borders.THIN  # 上
        border.right = xlwt.Borders.THIN  # 右
        border.bottom = xlwt.Borders.THIN  # 下
        border.left_colour = 0x40  # 边框线颜色
        border.right_colour = 0x40
        border.top_colour = 0x40
        border.bottom_colour = 0x40
        style.borders = border

        return style

    def main(self, needReadStr):
        try:
            print("开始生成报告，请等待...")
            text = needReadStr.split("<!-- -*- -*- split -*- -*- -->")

            htmldemo = text[0].replace('<br>', '').replace('<br />', '').replace('&nbsp;', ' ')
            regex = re.compile(r"\s{2}\s+")
            htmldemo = regex.sub('\n', htmldemo)

            soup = BeautifulSoup(htmldemo, 'lxml')
            ahrefs = soup.find_all('tbody')
            trs = ahrefs[0].find_all('tr')

            OALists = []
            applyOAList = []
            for i in range(0, len(trs)):
                tds = trs[i].find_all('td')
                json_data = {'oa': 0, 'title': '', 'acTime': '', 'catagory': '', 'status': '', 'scoreA': '', 'scoreB': ''}
                for j in range(1, len(tds)-1):
                    if j == 1:
                        try:
                            json_data['oa'] = re.match(r'/public/activity/activityDetail\.action\?activityId=(.*)', tds[j].find('a').get('href')).group(1).replace('\n', '').replace('\t', '').replace('\r', '')
                            OALists.append(json_data['oa'].replace('\n', '').replace('\t', '').replace('\r', ''))
                        except:
                            json_data['oa'] = 1000100
                            OALists.append('1000100')
                        json_data['title'] = tds[j].get_text().replace('\n', '').replace('\t', '').replace('\r', '')
                    elif j == 2:
                        json_data['catagory'] = tds[j].get_text().replace('\n', '').replace('\t', '').replace('\r', '')
                    elif j == 3:
                        json_data['acTime'] = tds[j].get_text().replace('\n', '').replace('\t', '').replace('\r', '')
                    elif j == 4:
                        json_data['status'] = tds[j].get_text().replace('\n', '').replace('\t', '').replace('\r', '')
                applyOAList.append(json_data)

            htmldemo = text[1].replace('<br>', '').replace(
                '<br />', '').replace('&nbsp;', ' ')
            regex = re.compile(r"\s{2}\s+")
            htmldemo = regex.sub('\n', htmldemo)

            soup = BeautifulSoup(htmldemo, 'lxml')
            ahrefs = soup.find_all('tbody')

            trs = ahrefs[0].find_all('tr')
            person_info = {'usercode': '', 'username': ''}
            for i in range(0, len(trs)):
                tds = trs[i].find_all('td')
                for j in range(0, len(tds)):
                    if "学号：" in tds[j].text:
                        person_info['usercode'] = tds[j].text.replace('学号：', '')
                    elif "姓名：" in tds[j].text:
                        person_info['username'] = tds[j].text.replace('姓名：', '')

            trs = ahrefs[1].find_all('tr')
            for i in range(0, len(trs)):
                tds = trs[i].find_all('td')
                knowWhere = OALists.index(tds[2].get_text().replace('\n', '').replace('\t', '').replace('\r', ''))
                scoreA = eval(tds[4].get_text().replace('\n', '').replace('\t', '').replace('\r', ''))
                scoreB = eval(tds[5].get_text().replace('\n', '').replace('\t', '').replace('\r', ''))
                if scoreA != 0:
                    applyOAList[knowWhere]['scoreA'] = scoreA
                if scoreB != 0:
                    applyOAList[knowWhere]['scoreB'] = scoreB
                applyOAList[knowWhere]['title'] = tds[0].get_text().replace('\n', '').replace('\t', '').replace('\r', '')

            wk = xlwt.Workbook()
            times = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            sheet2 = wk.add_sheet(f"总体信息", cell_overwrite_ok=True)
            sheet1 = wk.add_sheet(f"oaApplyList", cell_overwrite_ok=True)
            sheet1.write_merge(0, 0, 0, 6, f"活动申请结果及加分详单",self.centerStyle(True, 230))
            sheet1.write_merge(2, 2, 0, 6, f"报表生成时间：{times}", self.centerStyle())
            sheet1.write(1, 0, "学号", self.centerStyle())
            sheet1.write(1, 2, "姓名", self.centerStyle())
            sheet1.write_merge(1, 1, 1, 1, person_info['usercode'], self.centerStyle())
            sheet1.write_merge(1, 1, 3, 6, person_info['username'], self.centerStyle())
            sheet1.write(3, 0, "OA", self.centerStyle())
            sheet1.write(3, 1, "标题", self.centerStyle())
            sheet1.write(3, 2, "类型", self.centerStyle())
            sheet1.write(3, 3, "申请时间", self.centerStyle())
            sheet1.write(3, 4, "订单状态", self.centerStyle())
            sheet1.write(3, 5, "得分", self.centerStyle())
            sheet1.write(3, 6, "诚信分", self.centerStyle())
            sheet1.col(1).width = 330 * 30  # 定义列宽
            sheet1.col(2).width = 110 * 30  # 定义列宽
            sheet1.col(3).width = 165 * 30  # 定义列宽
            for i in range(0, len(applyOAList)):
                sheet1.write(i+4, 0, applyOAList[i]['oa'])
                sheet1.write(i+4, 1, applyOAList[i]['title'])
                sheet1.write(i+4, 2, applyOAList[i]['catagory'])
                sheet1.write(i+4, 3, applyOAList[i]['acTime'])
                sheet1.write(i+4, 4, applyOAList[i]['status'])
                sheet1.write(i+4, 5, applyOAList[i]['scoreA'])
                sheet1.write(i+4, 6, applyOAList[i]['scoreB'])

            sheet2.write_merge(0, 0, 0, 6, "第二课堂 活动与得分汇总信息",self.centerStyle(True, 250))
            sheet2.write_merge(2, 2, 0, 6, f"报表生成时间：{times}", self.centerStyle())
            sheet2.write(1, 0, "学号", self.centerStyle())
            sheet2.write(1, 3, "姓名", self.centerStyle())
            sheet2.write_merge(1, 1, 1, 2, person_info['usercode'], self.centerStyle())
            sheet2.write_merge(1, 1, 4, 6, person_info['username'], self.centerStyle())
            sheet2.write(3, 0, "类型", self.centerStyle())
            sheet2.write(4, 0, "志愿", self.centerStyle())
            sheet2.write(5, 0, "三创", self.centerStyle())
            sheet2.write(6, 0, "讲座", self.centerStyle())
            sheet2.write(7, 0, "校园文化", self.centerStyle())
            sheet2.write(8, 0, "实践", self.centerStyle())
            sheet2.write(9, 0, "安全文明", self.centerStyle())
            sheet2.write(3, 1, "得分A", self.centerStyle())
            sheet2.write(3, 2, "得分B", self.centerStyle())
            sheet2.write(4, 1, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"志愿公益",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(4, 2, "", self.get_style())
            sheet2.write(5, 1, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"创新创业创意",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(5, 2, "", self.get_style())
            sheet2.write(6, 1, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"讲座报告",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(6, 2, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"主题教育",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(7, 1, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"校园文化活动",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(7, 2, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"校园文化竞赛活动",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(8, 1, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"社会实践",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(8, 2, "", self.get_style())
            sheet2.write(9, 1, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"校园文明",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(9, 2, xlwt.Formula(f'SUMIF(oaApplyList!$C$2:$C${len(applyOAList)+1},"安全教育网络教学",oaApplyList!$F$2:$F${len(applyOAList)+1})'), self.get_style())
            sheet2.write(3, 3, "总需得分", self.centerStyle())
            sheet2.write(4, 3, 1, self.get_style())
            sheet2.write(5, 3, 1.5, self.get_style())
            sheet2.write(6, 3, 1.5, self.get_style())
            sheet2.write(7, 3, 1, self.get_style())
            sheet2.write(8, 3, 2, self.get_style())
            sheet2.write(9, 3, 1, self.get_style())
            sheet2.write(3, 4, "当前学分", self.centerStyle())
            sheet2.write(4, 4, xlwt.Formula('IF(B5+C5>1,1,B5+C5)'), self.get_style())
            sheet2.write(5, 4, xlwt.Formula('IF(B6+C6>1.5,1.5,B6+C6)'), self.get_style())
            sheet2.write(6, 4, xlwt.Formula('IF(B7+C7>1.5,1.5,B7+C7)'), self.get_style())
            sheet2.write(7, 4, xlwt.Formula('IF(B8+C8>1,1,B8+C8)'), self.get_style())
            sheet2.write(8, 4, xlwt.Formula('IF(B9+C9>2,2,B9+C9)'), self.get_style())
            sheet2.write(9, 4, xlwt.Formula('IF(B10+C10>1,1,B10+C10)'), self.get_style())
            sheet2.write(3, 5, "共得分", self.centerStyle())
            sheet2.write(4, 5, xlwt.Formula('B5+C5'), self.get_style())
            sheet2.write(5, 5, xlwt.Formula('B6+C6'), self.get_style())
            sheet2.write(6, 5, xlwt.Formula('B7+C7'), self.get_style())
            sheet2.write(7, 5, xlwt.Formula('B8+C8'), self.get_style())
            sheet2.write(8, 5, xlwt.Formula('B9+C9'), self.get_style())
            sheet2.write(9, 5, xlwt.Formula('B10+C10'), self.get_style())
            sheet2.write(3, 6, "剩余需得分", self.centerStyle())
            sheet2.col(6).width = 100 * 30  # 定义列宽
            sheet2.write(4, 6, xlwt.Formula('IF(B5+C5-D5>0,0,-(B5+C5-D5))'), self.get_style())
            sheet2.write(5, 6, xlwt.Formula('IF(B6+C6-D6>0,0,-(B6+C6-D6))'), self.get_style())
            sheet2.write(6, 6, xlwt.Formula('IF(B7+C7-D7>0,0,-(B7+C7-D7))'), self.get_style())
            sheet2.write(7, 6, xlwt.Formula('IF(B8+C8-D8>0,0,-(B8+C8-D8))'), self.get_style())
            sheet2.write(8, 6, xlwt.Formula('IF(B9+C9-D9>0,0,-(B9+C9-D9))'), self.get_style())
            sheet2.write(9, 6, xlwt.Formula('IF(B10+C10-D10>0,0,-(B10+C10-D10))'), self.get_style())
            sheet2.write(10, 0, "", self.get_style())
            sheet2.write(10, 1, "", self.get_style())
            sheet2.write(10, 2, "", self.get_style())
            sheet2.write(10, 3, xlwt.Formula('SUM(D5:D10)'), self.get_style())
            sheet2.write(10, 4, xlwt.Formula('SUM(E5:E10)'), self.get_style())
            sheet2.write(10, 5, xlwt.Formula('SUM(F5:F10)'), self.get_style())
            sheet2.write(10, 6, xlwt.Formula('SUM(G5:G10)'), self.get_style())
            sheet2.write(1, 8, "注：")
            sheet2.write(2, 8, "1.本报表由系统自动生成，数据仅供参考。")
            sheet2.write(3, 8, "2.报表中的“总需得分”按照以下标准计算：主题报告1.5分、社会实践2分、创新创业创意1.5分、校园安全文明1分、志愿公益1分、校园文化1分，共计8学分。")
            sheet2.write(4, 8, "3.报表中的部分值若为空，请启用编辑功能后即可看到对应数据。")
            sheet2.write(1, 8, "注：")
            sheet2.write(2, 8, "1.OA号为1000100的代表无法获得该活动的OA号信息。")
            saveFileName = f"{person_info['usercode']}-{time.strftime('%Y%m%d%H%M%S')}-{str(uuid.uuid1())[:8] + str(uuid.uuid4())[9:13] + str(uuid.uuid4())[14:18] + str(uuid.uuid1())[19:23] + str(uuid.uuid4())[-12:]}.xls"
            wk.save(f"./{saveFileName}")
            return {"filename": saveFileName}
        except Exception as e:
            traceback.print_exc()
            return {"filename": "err", "reason": "捕获到了意料之外的异常，如有疑问可附带报错信息提交issue。"}


def main():
    try:
        print("本程序支持的浏览器有：\n1. Edge（推荐）\n2. Chrome\n3. Firefox")
        while True:
            try:
                browser = int(input("请选择浏览器(1-3)："))
                if browser in [1, 2, 3]:
                    break
                else:
                    print("输入错误，请重新输入！")
            except:
                print("输入错误，请重新输入！")

        print("正在启动浏览器，请稍候……\n(若先前未安装必要的前置程序，在此将自动下载并安装，请等待安装完成)\n=====================\n启动浏览器后，在报告未完成生成操作前，请勿关闭命令行及缩小浏览器窗口。")
        if browser == 1:
            options = webdriver.EdgeOptions()
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument('log-level=3')
            driver = webdriver.Edge(options=options, service=Service(EdgeChromiumDriverManager().install()))
        elif browser == 2:
            options = webdriver.ChromeOptions()
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument('log-level=3')
            driver = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))
        elif browser == 3:
            driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
        elif browser == 4:
            # opera
            # 虽然此处编写了相关代码，但实际运行会报错，原因不明。因此在接收用户输入时，此选项被禁用。
            # 如有需要，可手动修改代码并运行。
            # driver = webdriver.Opera(OperaDriverManager().install())
            pass
        elif browser == 5:
            # ie
            # 虽然此处编写了相关代码，但实际运行会存在不同程度的命令行卡死情况（win10/win11均已测试），原因不明。因此在接收用户输入时，此选项被禁用。
            # 如有需要，可手动修改代码并运行。
            # options = webdriver.IeOptions()
            # options.ignore_zoom_level = True
            # options.add_argument('log-level=3')
            # driver = webdriver.Ie(options=options, service=Service(IEDriverManager().install()))
            pass

        print("浏览器已启动，请在弹出的浏览器中输入OA账号密码登录。完成后按回车键继续。")
        driver.maximize_window()
        driver.get("https://authserver.sit.edu.cn/authserver/login?service=http%3A%2F%2Fsc.sit.edu.cn%2F")
        input(">>>按回车键继续>>>")
        driver.get("http://sc.sit.edu.cn/public/pcenter/activityOrderList.action?pageNo=1&pageSize=1000")
        content = driver.page_source + "\n<!-- -*- -*- split -*- -*- -->\n" 
        driver.get("http://sc.sit.edu.cn/public/pcenter/scoreDetail.action?pageNo=1&pageSize=1000")
        content = content + driver.page_source
        print("数据获取完成，即将关闭浏览器。")
        driver.quit()
        data = xls_generate().main(content)
        if data["filename"] == "err":
            print("生成失败。原因："+data["reason"])
        else:
            print("生成成功。文件名："+data["filename"]+"。已保存在本程序所在目录下。")
    except:
        print("捕获到了意料之外的异常，如有疑问可附带报错信息提交issue。")
        traceback.print_exc()
    input("按回车键可退出程序。")

if __name__ == '__main__':
    main()
