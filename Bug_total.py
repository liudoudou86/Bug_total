#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:lz

import os
import re
import time

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pymysql
import PySimpleGUI as sg
import xlwt


class CS():

    def zm(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_zm = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status IN ('新提交','待反测的') AND NAME LIKE '%"+project+"%' \
            AND creation_ts BETWEEN '"+start+"' AND '"+end+"' AND bug_severity = '致命';"
        # print(sql_zm)
        # 执行sql语句
        try:
            cursor.execute(sql_zm)
            result = cursor.fetchone()
            for bug in result:
                # print("致命:",bug,"个")
                return bug
        except:
            print("查询异常_致命")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def yz(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_yz = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status IN ('新提交','待反测的') AND NAME LIKE '%"+project+"%' \
            AND creation_ts BETWEEN '"+start+"' AND '"+end+"' AND bug_severity = '严重';"
        # print(sql_yz)
        # 执行sql语句
        try:
            cursor.execute(sql_yz)
            result = cursor.fetchone()
            for bug in result:
                # print("严重:",bug,"个")
                return bug
        except:
            print("查询异常_严重")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def yb(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_yb = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status IN ('新提交','待反测的') AND NAME LIKE '%"+project+"%' \
            AND creation_ts BETWEEN '"+start+"' AND '"+end+"' AND bug_severity = '一般';"
        # print(sql_yb)
        # 执行sql语句
        try:
            cursor.execute(sql_yb)
            result = cursor.fetchone()
            for bug in result:
                # print("一般:",bug,"个")
                return bug
        except:
            print("查询异常_一般")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()  

    def qw(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_qw = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status IN ('新提交','待反测的') AND NAME LIKE '%"+project+"%' \
            AND creation_ts BETWEEN '"+start+"' AND '"+end+"' AND bug_severity = '轻微';"
        # print(sql_qw)
        # 执行sql语句
        try:
            cursor.execute(sql_qw)
            result = cursor.fetchone()
            for bug in result:
                # print("轻微:",bug,"个")
                return bug
        except:
            print("查询异常_轻微")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def yyz(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_yyz = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status = '已验证' AND NAME LIKE '%"+project+"%' \
            AND lastdiffed BETWEEN '"+start+"' AND '"+end+"';"
        # print(sql_qw)
        # 执行sql语句
        try:
            cursor.execute(sql_yyz)
            result = cursor.fetchone()
            for bug in result:
                # print("今日已验证:",bug,"个")
                return bug
        except:
            print("查询异常_已验证")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def dfc(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_dfc = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status = '待反测的' AND NAME LIKE '%"+project+"%';"
        # print(sql_dfc)
        # 执行sql语句
        try:
            cursor.execute(sql_dfc)
            result = cursor.fetchone()
            for bug in result:
                # print("待反测共计:",bug,"个")
                return bug
        except:
            print("查询异常_待反测")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def xgz(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_xgz = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status IN ('新提交','问题未解决','修改中') AND NAME LIKE '%"+project+"%';"
        # print(sql_xgz)
        # 执行sql语句
        try:
            cursor.execute(sql_xgz)
            result = cursor.fetchone()
            for bug in result:
                # print("修改中共计:",bug,"个")
                return bug
        except:
            print("查询异常_修改中")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def bug(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_excel = "SELECT bugs.bug_id AS 问题ID,bugs.bug_severity AS 问题等级,components.name AS 问题模块,short_desc AS 问题描述,\
        profiles.realname AS 提交者 FROM bugs INNER JOIN products ON bugs.product_id = products.id INNER JOIN profiles ON bugs.reporter = profiles.userid\
        INNER JOIN components ON bugs.component_id = components.id WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的')\
        AND bugs.creation_ts BETWEEN '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命', '严重', '一般') ORDER BY bugs.bug_severity DESC;"
        # print(sql_excel)
        # 执行sql语句
        cursor.execute(sql_excel)
        fileds = [filed[0] for filed in cursor.description]
        all_date = cursor.fetchall()
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()
        try:
            #写入excel
            book = xlwt.Workbook() #创建一个book
            sheet = book.add_sheet('问题列表') #创建一个sheet表
            for col,filed in enumerate(fileds):
                sheet.write(0,col,filed)
            #从第一行开始写
            row = 1
            for data in all_date:
                for col,filed in enumerate(data):
                    sheet.write(row,col,filed)
                row += 1
            book.save('D:\问题列表.xls')
        except:
            print('导出excel异常')

class SJ():

    def analysis(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()

        # 图1的X轴
        sql_tu1_x = "SELECT components.name FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' GROUP BY components.name ORDER BY count(bugs.bug_severity) DESC;"
        # print(sql_tu1_x)
        cursor.execute(sql_tu1_x)
        result = cursor.fetchall()
        x_axis_1 = [x[0] for x in result]
        # print(x_axis_1)
        # 图1的Y轴
        sql_tu1_y = "SELECT count(bugs.bug_severity) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' GROUP BY components.name ORDER BY count(bugs.bug_severity) DESC;"
        # print(sql_tu1_y)
        cursor.execute(sql_tu1_y)
        result = cursor.fetchall()
        y_axis_1 = [x[0] for x in result]

        # 图2的X轴
        sql_tu2_x = "SELECT components.name FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的') AND bugs.creation_ts BETWEEN \
            '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命','严重','一般') GROUP BY components.name ORDER BY count(bugs.bug_severity) DESC;"
        # print(sql_tu2_x)
        cursor.execute(sql_tu2_x)
        result = cursor.fetchall()
        x_axis_2 = [x[0] for x in result]
        # print(x_axis_2)
        # 图2的Y轴
        sql_tu2_y = "SELECT count(bugs.bug_severity) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的') AND bugs.creation_ts BETWEEN \
            '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命','严重','一般') GROUP BY components.name ORDER BY count(bugs.bug_severity) DESC;"
        # print(sql_tu2_y)
        cursor.execute(sql_tu2_y)
        result = cursor.fetchall()
        y_axis_2 = [x[0] for x in result]

        # 图3的X轴
        sql_tu3_x = "SELECT bugs.cf_tijiaojieduan FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' GROUP BY bugs.cf_tijiaojieduan;"
        # print(sql_tu3_x)
        cursor.execute(sql_tu3_x)
        result = cursor.fetchall()
        x_axis_3 = [x[0] for x in result]
        # print(x_axis_3)
        # 图3的Y轴
        sql_tu3_y = "SELECT count(bugs.bug_severity) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' GROUP BY bugs.cf_tijiaojieduan;"
        # print(sql_tu3_y)
        cursor.execute(sql_tu3_y)
        result = cursor.fetchall()
        y_axis_3 = [x[0] for x in result]

        # 图4的X轴
        sql_tu4_x = "SELECT profiles.realname FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的') AND bugs.creation_ts BETWEEN \
            '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命','严重','一般') GROUP BY profiles.realname ORDER BY count(bugs.bug_severity) DESC;"
        # print(sql_tu4_x)
        cursor.execute(sql_tu4_x)
        result = cursor.fetchall()
        x_axis_4 = [x[0] for x in result]
        # print(x_axis_4)
        # 图4的Y轴
        sql_tu4_y = "SELECT count(bugs.bug_severity) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的') AND bugs.creation_ts BETWEEN \
            '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命','严重','一般') GROUP BY profiles.realname ORDER BY count(bugs.bug_severity) DESC;"
        # print(sql_tu4_y)
        cursor.execute(sql_tu4_y)
        result = cursor.fetchall()
        y_axis_4 = [x[0] for x in result]

        # 关闭数据库
        cursor.close()
        conn.close()

        mpl.rcParams["font.sans-serif"] = ["SimHei"] # 用黑体显示中文
        # mpl.rcParams["axes.unicode_minus"] = False
        plt.figure(figsize=(8,7), dpi=90)
        plt.figure(1)
        bar_width=0.25

        # 图1柱形图
        ax1 = plt.subplot(221)
        plt.bar(x = x_axis_1, height = y_axis_1, width=bar_width, color = '#1f77b4', alpha=0.8, label='-')
        # 在柱状图上显示具体数值，ha参数控制水平对齐方式，va参数控制垂直对齐方式
        for x, y in enumerate(y_axis_1):
            plt.text(x, y, '%s' % y, ha='center', va='bottom', fontsize=10, rotation=0)
        # 标题
        plt.title('项目问题模块趋势图') 
        # XY轴标题
        # plt.xlabel("问题模块")      
        # plt.ylabel('问题数量')

        # 图2柱形图
        ax2 = plt.subplot(222)
        plt.bar(x = x_axis_2, height = y_axis_2, width=bar_width, color = '#d62728', alpha=0.8, label='-')
        # 在柱状图上显示具体数值，ha参数控制水平对齐方式，va参数控制垂直对齐方式
        for x, y in enumerate(y_axis_2):
            plt.text(x, y, '%s' % y, ha='center', va='bottom', fontsize=10, rotation=0)
        # 标题
        plt.title('日别问题提交模块趋势图') 
        # XY轴标题
        # plt.xlabel("问题模块")      
        # plt.ylabel('问题数量')

        # 图3柱形图
        ax3 = plt.subplot(223)
        plt.bar(x = x_axis_3, height = y_axis_3, width=bar_width, color = '#2ca02c', alpha=0.8, label='-')
        # 在柱状图上显示具体数值，ha参数控制水平对齐方式，va参数控制垂直对齐方式
        for x, y in enumerate(y_axis_3):
            plt.text(x, y, '%s' % y, ha='center', va='bottom', fontsize=10, rotation=0)
        # 标题
        plt.title('项目问题阶段趋势图') 
        # XY轴标题
        # plt.xlabel("问题模块")      
        # plt.ylabel('问题数量')

        # 图4柱形图
        ax4 = plt.subplot(224)
        plt.bar(x = x_axis_4, height = y_axis_4, width=bar_width, color = '#ff7f0e', alpha=0.8, label='')
        # 在柱状图上显示具体数值，ha参数控制水平对齐方式，va参数控制垂直对齐方式
        for x, y in enumerate(y_axis_4):
            plt.text(x, y, '%s' % y, ha='center', va='bottom', fontsize=10, rotation=0)
        # 标题
        plt.title('日别问题提交人员趋势图') 
        # XY轴标题
        # plt.xlabel("问题模块")      
        # plt.ylabel('问题数量')

        # 显示图例
        plt.legend()
        plt.show()

start_time = str((time.strftime('%Y-%m-%d 00:00:00',time.localtime())))
end_time = str((time.strftime('%Y-%m-%d 23:59:59',time.localtime())))

layout = [
    [sg.Radio('测试机', 'RADIO1', key='_RADIO1_', default=True), sg.Radio('工作机', 'RADIO1', key='_RADIO2_')],
    [sg.Text('项目名称: ',font='微软雅黑',size=(10, 1)),sg.Input(key='_PROJECT_')],
    [sg.Text('开始时间: ',font='微软雅黑',size=(10, 1)),sg.InputText(start_time,key='_START_')], #默认初始内容的输入框
    [sg.Text('结束时间: ',font='微软雅黑',size=(10, 1)),sg.InputText(end_time,key='_END_')], #默认初始内容的输入框
    [sg.Text('致命: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_zm_',size=(10, 1))],
    [sg.Text('严重: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_yz_',size=(10, 1))],
    [sg.Text('一般: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_yb_',size=(10, 1))],
    [sg.Text('轻微:',font='微软雅黑',size=(10, 1)), sg.Text('', key='_qw_',size=(10, 1))],
    [sg.Text('今日共提交: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_ttl_',size=(10, 1))],
    [sg.Text('今日已验证: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_yyz_',size=(10, 1))],
    [sg.Text('待反测共计: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_dfc_',size=(10, 1))],
    [sg.Text('修改中共计: ',font='微软雅黑',size=(10, 1)), sg.Text('', key='_xgz_',size=(10, 1))],
    [sg.Button('确认',key = '_CONFIRM_',font='微软雅黑', size=(10, 1)), sg.Exit('退出',key = '_EXIT_',font='微软雅黑', size=(10, 1)), sg.Open('打开Bug列表',key = '_OPEN_',font='微软雅黑', size=(10, 1))]
]
# 定义窗口，窗口名称
window = sg.Window('Bug数量统计工具',layout,font='微软雅黑')
# 自定义窗口进行数值回显
while True:
    event,values = window.read()
    if event == '_CONFIRM_':
        if values.get('_RADIO1_','True'):
            info = str('192.168.25.247')
        elif values.get('_RADIO2_','True'):
            info = str('10.30.12.247')
        else:
            pass
        project = str(values['_PROJECT_'])
        # print(project)
        start = str(values['_START_'])
        # print(start)
        end = str(values['_END_'])
        # print(end)
        zm = CS().zm()
        yz = CS().yz()
        yb = CS().yb()
        qw = CS().qw()
        ttl = zm + yz + yb + qw
        yyz = CS().yyz()
        dfc = CS().dfc()
        xqz = CS().xgz()
        window['_zm_'].Update(zm) #将读取的数据回显到界面
        window['_yz_'].Update(yz)
        window['_yb_'].Update(yb)
        window['_qw_'].Update(qw)
        window['_ttl_'].Update(ttl)
        window['_yyz_'].Update(yyz)
        window['_dfc_'].Update(dfc)
        window['_xgz_'].Update(xqz)
        CS().bug()
        SJ().analysis()
    elif event == '_OPEN_':
        os.system(r"D:\问题列表.xls")
    elif event in ['_EXIT_',None]:
        break
    else:
        pass
