# -*- codeing = utf-8 -*-
# @Time : 2022/6/1 12:30
# @Author : Administrator
# @File : top250.py
# @Software: PyCharm
#爬取豆瓣top250电影 所有信息

from requests_html import HTMLSession   #直接抓取网页
import openpyxl     #excel
import re       #正则
import os
import time
from openpyxl.styles import Alignment,Border,Side

#获取信息
def info_find():
    i = 0
    list_250 = []   #存储250条数据

    while True:
        url = f'https://movie.douban.com/top250?start={i}&filter='
        info = session.get(url)  # get直接请求整个网址

        try:
            for j in range(1,26):
                sel = f'#content > div > div.article > ol > li:nth-child({j})'  # selector
                # content > div > div.article > ol > li:nth-child(1)
                content = info.html.find(sel)  # 定位selector

                web = list(content[0].absolute_links)[0]    #电影链接
                info_text = content[0].text + '\n' + web  # 返回是列表所以要下标

                list_250.append(info_text)
        except:
            # 会出现下标越界，其实已经获取完了
            pass

        i = i + 25
        # 控制循环结束
        if i == 250:
            break

    print('获取完成')

    return list_250

#解析内容
def analysis_content(content_text):
    #后续用来存储excel
    excel_list = []

    for k in content_text:

        dict_all = {
            '电影名称': '',
            '导演': '',
            '部分主演': '',
            '电影上映年份': '',
            '电影上映城市': '',
            '电影类别': '',
            '评分': '',
            '评价人数': '',
            '简介': '',
            '是否可播放': '',
            '电影介绍网址': '',
        }

        dict_all['电影名称'] = k.split('\n')[1].replace('\xa0', '')
        if '[可播放]' in dict_all['电影名称']:
            dict_all['是否可播放'] = '可播放'
        else:
            dict_all['是否可播放'] = '暂不可播放'
        dict_all['电影名称'] = dict_all['电影名称'].replace('[可播放]', '').strip()

        #页面显示不全，有时没有主演显示
        try:
            # 导演
            director = re.findall(r'(?<=导演:)[\S\s]*(?=主演:)', k.split('\n')[2])[0].replace('\xa0', '').strip()
            # 部分主演
            lead = re.findall(r'(?<=主演:)[\S\s]*', k.split('\n')[2])[0].replace('\xa0', '').strip()
            dict_all['导演'] = director
            dict_all['部分主演'] = lead
        except:
            director = re.findall(r'(?<=导演:)[\S\s]*', k.split('\n')[2])[0].replace('\xa0', '').strip()
            dict_all['导演'] = director
            dict_all['部分主演'] = '页面无具体显示,可前往电影介绍网址查看'

        dict_all['电影上映年份'] = k.split('\n')[3].split('/')[0].replace('\xa0', '').strip()
        dict_all['电影上映城市'] = k.split('\n')[3].split('/')[1].replace('\xa0', '').strip()
        dict_all['电影类别'] = k.split('\n')[3].split('/')[2].replace('\xa0', '').strip()

        dict_all['评分'] = k.split('\n')[4].split(' ')[0]
        dict_all['评价人数'] = k.split('\n')[4].split(' ')[1]

        #有的没简介，需要处理
        dict_all['简介'] = k.split('\n')[5]
        try:
            dict_all['电影介绍网址'] = k.split('\n')[6]
        except:
            dict_all['简介'] = '暂无简介'
            dict_all['电影介绍网址'] = k.split('\n')[5]

        # print(dict_all)
        excel_list.append(dict_all)

    print('解析完成')
    # print(excel_list)
    return excel_list

#插入excel中
def insert_excel(excel_info,excel_savepath):
    # excel保存位置     time.strftime("%Y-%m-%d+%H.%M.%S")记录每一次的时间
    outpath = excel_savepath + "\\" + '豆瓣top250电影信息' + '-' + time.strftime("%Y-%m-%d+%H.%M.%S") + ".xlsx"
    #excel模板位置
    excel_model = excel_savepath + "\\" + '新建 Microsoft Excel 工作表.xlsx'

    outwb = openpyxl.load_workbook(excel_model)
    workSheet = outwb['Sheet1']     #激活

    num = 2
    for info in excel_info:     #遍历每一个字典集
        els = 1
        for key,values in info.items():
            # print(values)
            workSheet.cell(num, els).value = values
            els = els + 1
        # print('下一条')
        num = num + 1

    #调整每一列的列宽
    workSheet.column_dimensions['A'].width = 40
    workSheet.column_dimensions['B'].width = 40
    workSheet.column_dimensions['C'].width = 40
    workSheet.column_dimensions['D'].width = 16
    workSheet.column_dimensions['E'].width = 17
    workSheet.column_dimensions['F'].width = 17
    workSheet.column_dimensions['G'].width = 7
    workSheet.column_dimensions['H'].width = 15
    workSheet.column_dimensions['I'].width = 25
    workSheet.column_dimensions['J'].width = 16
    workSheet.column_dimensions['K'].width = 43

    #行高
    for hei in range(2,len(excel_info)+2):
        workSheet.row_dimensions[hei].height = 50

    #居中
    alignment_center = Alignment(horizontal='center', wrap_text=True, vertical='center')  #, wrap_text=True
    ws_area = workSheet["A2:K251"]   #如字段变多需重新设计
    for l in ws_area:
        for j in l:
            j.alignment = alignment_center

    #边框
    side = Side('thin') #细线
    content_Border = Border(left=side,right=side,top=side,bottom=side)
    ws_area = workSheet["A1:K251"]  # 如字段变多需重新设计
    for l in ws_area:
        for j in l:
            j.border = content_Border

    outwb.save(outpath)
    pass

if __name__ == '__main__':
    retval = os.getcwd()    #获取当前工作目录
    print("当前工作目录为 %s" % retval)

    # 建立基础对话 让Python作为一个客户端，和远端服务器交谈
    session = HTMLSession()

    print('-' * 20)
    print('正在获取内容')
    reinfo = info_find()

    print('-' * 20)
    print('正在解析内容')
    excel = analysis_content(reinfo)

    print('-' * 20)
    print('存入EXCEL中')
    insert_excel(excel,retval)
    print('存入成功')




