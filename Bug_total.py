#!/usr/bin/env python
# -*- coding:utf-8 -*-
# Author:lz

import re
import os
import pymysql
import time
import xlwt
import matplotlib.pyplot as plt
import matplotlib as mpl
import numpy as np


def gateway():
    # 粗略找IP地址的正则表达式，包括掩码在内一起找
    pr = re.compile(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}')
    # 通过默认路由得到网关地址，可能存在多条，但只取第一条[0]，通过正则找IP得到的列表第三个元素就是网关[2]
    gw = re.findall(pr,os.popen('route print | find " 0.0.0.0 "').read().strip().split('\n')[0])[2]
    # 向网关发udp请求包，避免无脑的随意udp包游荡在内网中，通过udp的ip头部获取到发包的源地址，也就是本机生产地址
    wg = (gw[:-5])
    if wg == '192.168':
        ip_cs = str('192.168.25.247')
        return ip_cs
    else:
        ip_gz = str('10.30.12.247')
        return ip_gz

info = gateway()
# print(info)

project = str(input("请输入项目名称："))
start = str(time.strftime('%Y-%m-%d 00:00:00',time.localtime()))
end = str(time.strftime('%Y-%m-%d 23:59:59',time.localtime()))


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
                print("致命:",bug,"个")
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
                print("严重:",bug,"个")
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
                print("一般:",bug,"个")
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
                print("轻微:",bug,"个")
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
                print("今日已验证:",bug,"个")
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
                print("待反测共计:",bug,"个")
            return bug
        except:
            print("查询异常_待反测")
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()

    def qt(self):
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # sql语句
        sql_qt = "SELECT count(*) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            WHERE bug_status IN ('新提交','问题未解决','修改中') AND NAME LIKE '%"+project+"%';"
        # print(sql_qt)
        # 执行sql语句
        try:
            cursor.execute(sql_qt)
            result = cursor.fetchone()
            for bug in result:
                print("修改中共计:",bug,"个")
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
        AND bugs.creation_ts BETWEEN '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命', '严重', '一般');"
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
            book.save('今日Bug统计.xls')
        except:
            print('导出excel异常')

class SJ():

    def analysis(self):
        # 连接到mysql数据库
        # TODO：项目：XM20200255-纪检监察办案指挥系统V3.1T定版集成
        # 连接到mysql数据库
        conn = pymysql.connect(host=info, user='report', password='report', port=3306, db='bugs', charset='UTF8')
        # 创建游标对象
        cursor = conn.cursor()
        # X轴的问题轴线图
        sql_mokuai = "SELECT components.name FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的') AND bugs.creation_ts BETWEEN \
            '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命', '严重', '一般') GROUP BY components.name ORDER BY count(bugs.bug_severity)DESC;"
        # print(sql_analysis)
        # 执行sql语句
        cursor.execute(sql_mokuai)
        result = cursor.fetchall()
        x_axis = [x[0] for x in result]
        # print(x_axis)
        # Y轴的问题轴线图
        sql_zhiming = "SELECT count(bugs.bug_severity) FROM bugs INNER JOIN products ON bugs.product_id = products.id \
            INNER JOIN profiles ON bugs.reporter = profiles.userid INNER JOIN components ON bugs.component_id = components.id \
            WHERE products.name LIKE '%"+project+"%' AND bugs.bug_status IN ('新提交', '待反测的') AND bugs.creation_ts BETWEEN \
            '"+start+"' AND '"+end+"' AND bugs.bug_severity IN ('致命','严重','一般') GROUP BY components.name ORDER BY count(bugs.bug_severity)DESC;"
        # 执行sql语句
        cursor.execute(sql_zhiming)
        result = cursor.fetchall()
        y_axis = [x[0] for x in result]
        # print(y_axis)
        # 关闭游标
        cursor.close()
        # 关闭数据库
        conn.close()
        # TODO：项目：XM20200255-纪检监察办案指挥系统V3.1T定版集成

        mpl.rcParams["font.sans-serif"] = ["SimHei"]
        mpl.rcParams["axes.unicode_minus"] = False

        bar_width=0.35
        plt.bar(x = x_axis, height = y_axis, width=bar_width, color = ['b','g','r','c','m','y','k','b','g','r','c','m','y','k'], alpha=0.6, label='问题总数')
        # 在柱状图上显示具体数值，ha参数控制水平对齐方式，va参数控制垂直对齐方式
        for x, y in enumerate(y_axis):
            plt.text(x, y, '%s' % y, ha='center', va='bottom')
        # 标题
        plt.title('日别问题发掘趋势图') 
        # XY轴标题
        plt.xlabel("问题模块")      
        plt.ylabel('问题数量')
        # 显示图例
        plt.legend()
        plt.show()


zm = CS().zm()
yz = CS().yz()
yb = CS().yb()
qw = CS().qw()
ttl = zm + yz + yb + qw
print('今日共提交: ',ttl,'个')
yyz = CS().yyz()
CS().dfc()
CS().qt()
CS().bug()
SJ().analysis()
os.system("pause")