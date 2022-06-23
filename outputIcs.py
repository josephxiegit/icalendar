# coding: utf-8
# !/usr/bin/python

import sys
import time, datetime
from random import Random
import json
import importlib
import os
import sqlite3
import pandas as pd

__author__ = 'joseph'

checkFirstWeekDate = 0
checkReminder = 1

YES = 0
NO = 1

DONE_firstWeekDate = time.time()
DONE_reminder = ""
DONE_EventUID = ""
DONE_UnitUID = ""
DONE_CreatedTime = ""
DONE_ALARMUID = ""
parent_os = os.path.dirname(os.path.realpath(__file__))
conf_classInfo_os = parent_os + "/conf_classInfo.json"
conf_classTime_os = parent_os + "/conf_classTime.json"

classTimeList = []
classInfoList = []


def main():
    basicSetting()
    uniteSetting()
    list_ClassInfoList = classInfoHandle()
    icsCreateAndSave(list_ClassInfoList)


def save(string):
    # classics_os = parent_os + "/class.ics"

    # 利用时间戳定义class.ics文件名称
    # time_stamp = str(int(time.time()))
    # today_name = str(datetime.date.today())
    # ics_name = 'class_' + today_name + '_' + time_stamp + '.ics'

    parent = os.path.dirname(os.path.realpath(__file__))
    classics_os = parent + "/classIcalendar.ics"

    f = open(classics_os, 'wb')
    f.write(string.encode("utf-8"))
    f.close()

    print("saved as path %s\n" % (classics_os))

def icsCreateAndSave(list_ClassInfoList):  # 写成ics文件并保存ics
    icsString = "BEGIN:VCALENDAR\nCALSCALE:GREGORIAN\nVERSION:2.0\nX-WR-CALNAME:课程表\nPRODID:-//Apple Inc.//Mac OS X 10.12//EN\nX-APPLE-CALENDAR-COLOR:#FC4208\nX-WR-TIMEZONE:Asia/Shanghai\nCALSCALE:GREGORIAN\nBEGIN:VTIMEZONE\nTZID:Asia/Shanghai\nBEGIN:STANDARD\nTZOFFSETFROM:+0900\nRRULE:FREQ=YEARLY;UNTIL=19910914T150000Z;BYMONTH=9;BYDAY=3SU\nDTSTART:19890917T000000\nTZNAME:GMT+8\nTZOFFSETTO:+0800\nEND:STANDARD\nBEGIN:DAYLIGHT\nTZOFFSETFROM:+0800\nDTSTART:19910414T000000\nTZNAME:GMT+8\nTZOFFSETTO:+0900\nRDATE:19910414T000000\nEND:DAYLIGHT\nEND:VTIMEZONE\n"
    global classTimeList, DONE_ALARMUID, DONE_UnitUID
    eventString = ""
    # print('list_ClassInfoList_final:', list_ClassInfoList)
    print("df_final:\n", pd.DataFrame(list_ClassInfoList,
                                      columns=['classTime', 'date', 'startTime', 'endTime', 'className', 'CREATED',
                                               'DTSTAMP', 'UID']))
    # print('classTimeList', classTimeList)

    for classInfo in list_ClassInfoList:
        # print('classInfo:', classInfo)
        # i = int(classInfo["classTime"]-1)
        className = classInfo["className"]
        endTime = str(classInfo["endTime"])
        startTime = str(classInfo["startTime"])

        # endTime = classTimeList[i]["endTime"]
        # startTime = classTimeList[i]["startTime"]

        eventString = eventString + "BEGIN:VEVENT\nCREATED:" + str(classInfo["CREATED"])
        eventString = eventString + "\nUID:" + str(classInfo["UID"])
        eventString = eventString + "\nDTEND;TZID=Asia/Shanghai:" + str(classInfo["date"]) + "T" + endTime
        eventString = eventString + "00\nTRANSP:OPAQUE\nX-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC\nSUMMARY:" + str(
            className)
        eventString = eventString + "\nDTSTART;TZID=Asia/Shanghai:" + str(classInfo["date"]) + "T" + startTime + "00"
        eventString = eventString + "\nDTSTAMP:" + DONE_CreatedTime
        eventString = eventString + "\nSEQUENCE:0\nBEGIN:VALARM\nX-WR-ALARMUID:" + DONE_ALARMUID
        eventString = eventString + "\nUID:" + DONE_UnitUID
        eventString = eventString + "\nTRIGGER:" + DONE_reminder
        eventString = eventString + "\nDESCRIPTION:事件提醒\nACTION:DISPLAY\nEND:VALARM\nEND:VEVENT\n"

    icsString = icsString + eventString + "END:VCALENDAR"
    save(icsString)

def classInfoHandle():  # 处理class信息
    global classInfoList
    global DONE_firstWeekDate
    i = 0
    # print('classInfoList', classInfoList)
    for classInfo in classInfoList:
        # 设置 UID
        global DONE_CreatedTime, DONE_EventUID
        CreateTime()  # 建立 DONE_CreatedTime，
        classInfo["CREATED"] = DONE_CreatedTime
        classInfo["DTSTAMP"] = DONE_CreatedTime
        classInfo["UID"] = UID_Create()
    # print('classInfoList', classInfoList)

    # 生成ICS源数据
    list_classInfoListFullLessons = getClassInfoListFullLessons(classInfoList)
    # print('list_classInfoListFullLessons:\n', list_classInfoListFullLessons)

    df_classInfoList = listToDataFrame(list_classInfoListFullLessons)
    # print('df_classInfoList\n', df_classInfoList)

    df_ics = handlePdClassInfoList(df_classInfoList)

    # print('df_ics1:\n', df_ics)

    def uidDateNanListToString(df_database):
        # 将df中nan变为无
        list_database = df_database.values.tolist()
        for item in list_database:
            if str(item[4]) == "nan":
                index = list_database.index(item)
                df_database.loc[index, 'className'] = " "
                df_database.loc[index, 'CREATED'] = " "
                df_database.loc[index, 'DTSTAMP'] = " "
        return df_database

    df_ics = uidDateNanListToString(df_ics)
    # print('\ndf_ics:\n', df_ics)

    # 上传数据库
    to_database(df_ics)

    # 转化为字典
    list_columns = df_ics.columns.values.tolist()
    # print(list_columns)

    list_ClassInfoList = []
    for i in df_ics.values.tolist():
        dic_lesson = {
            list_columns[0]: '',
            list_columns[1]: '',
            list_columns[2]: '',
            list_columns[3]: '',
            list_columns[4]: '',
            list_columns[5]: '',
            list_columns[6]: '',
            list_columns[7]: '',
        }
        # print(i)
        for j in range(0, 6):
            if str(i[j]) == 'nan':
                i[j] = ' '

        dic_lesson['classTime'] = i[0]
        dic_lesson['date'] = i[1]
        dic_lesson['startTime'] = i[2]
        dic_lesson['endTime'] = i[3]
        dic_lesson['className'] = i[4]
        dic_lesson['CREATED'] = i[5]
        dic_lesson['DTSTAMP'] = i[6]
        dic_lesson['UID'] = i[7]

        list_ClassInfoList.append(dic_lesson)
    # print('list_ClassInfoList', list_ClassInfoList)
    return list_ClassInfoList


def getClassInfoListFullLessons(list):
    # 添加无课程数据
    # print('listIncompletedLessons:', list)
    list_date = []  # 得到list用不重复的date列表
    for lesson in list:
        if lesson['date'] not in list_date:
            list_date.append(lesson['date'])
    # print('list_date:\n', list_date)		# ['20210220', '20210221']

    listCompletedLessons = []  # 建立空的完整class列表
    for date in list_date:
        for classTime in range(1, 7):
            listCompletedLessons.append({'classTime': classTime, 'date': date})
    # print('listCompletedLessons:\n', listCompletedLessons)

    for lesson in list:  # 填充数据
        # print('lesson:\n', lesson)
        date = lesson['date']
        classTime = lesson['classTime']
        for i in listCompletedLessons:
            # print('i:\n', i)
            if i['classTime'] == classTime and i['date'] == date:
                # print('lesson', lesson)
                i['className'] = lesson['className']
                i['week'] = lesson['week']
                i['weekday'] = lesson['weekday']
                i['classroom'] = lesson['classroom']
                i['CREATED'] = lesson['CREATED']
                i['DTSTAMP'] = lesson['DTSTAMP']
                i['UID'] = lesson['UID']
                i['startTime'] = lesson['startTime']
                i['endTime'] = lesson['endTime']

    # print('listCompletedLessons:\n', listCompletedLessons)
    def getStartTime(classtime):
        if classtime == 1:
            return "0800"
        if classtime == 2:
            return "1000"
        if classtime == 3:
            return "1300"
        if classtime == 4:
            return "1500"
        if classtime == 5:
            return "1730"
        if classtime == 6:
            return "2000"

    def getEndTime(classtime):
        if classtime == 1:
            return "1000"
        if classtime == 2:
            return "1200"
        if classtime == 3:
            return "1500"
        if classtime == 4:
            return "1700"
        if classtime == 5:
            return "1930"
        if classtime == 6:
            return "2200"

    # 	i['startTime'] = getStartTime(i['classTime'])
    # 	i['endTime'] = getStartTime(i['classTime'])
    for lesson in listCompletedLessons:
        if len(lesson) == 2:
            # print(lesson)
            lesson['startTime'] = getStartTime(lesson['classTime'])
            lesson['endTime'] = getStartTime(lesson['classTime'])
    # print('listCompletedLessons:', listCompletedLessons)
    return listCompletedLessons

def listToDataFrame(list):
    # list转成pandas
    df = pd.DataFrame(list)
    # print('\ndf:\n', df)
    order = ['classTime', 'date', 'startTime', 'endTime', 'className', 'week', 'weekday', 'classroom', 'CREATED',
             'DTSTAMP', 'UID']
    df = df[order]
    # df.drop(df.columns[[3, 4, 5]], axis=1, inplace=True)
    df.drop(df.columns[[5, 6, 7]], axis=1, inplace=True)
    # print('\ndf:\n', df)
    return df


def handlePdClassInfoList(df):
    # 处理dfClassInfoList - 替换成数据库内的UID
    address_db = parent_os + '/dbcalender.sqlite3'
    print(address_db)
    conn = sqlite3.connect(address_db)
    list_ClassInfoList = df.values.tolist()
    # print('list_ClassInfoList', list_ClassInfoList)

    # 替换成数据库内的UID
    for lesson in list_ClassInfoList:
        classtime = str(lesson[0])
        date = str(lesson[1])
        sql_exist = "SELECT classname, uid, classTime, date FROM app01_icalendar WHERE classtime='" + classtime + "' and date='" + date + "'"
        df_read = pd.read_sql(sql_exist, conn)
        # print('df_read:', df_read)
        for i in df_read.values.tolist():
            # print('uid_database:', i[1])    # 查询数据库内的uid号
            if i[2] == classtime and i[3] == date:
                index = list_ClassInfoList.index(lesson)
                df.loc[index, 'UID'] = i[1]
    # print('\ndf:\n', df)
    # 给没有UID的课增加UID等
    for i in range(0, df.shape[0]):  # 行数
        for j in range(0, df.shape[1]):  # 列数
            if str(df.iloc[i, 7]) == 'nan':
                df.iloc[i, 5] = DONE_CreatedTime
                df.iloc[i, 6] = DONE_CreatedTime
                df.iloc[i, 7] = UID_Create()
    # print('\ndf:\n', df)
    return df


def to_database(df):
    # 添加数据库;将数据库中uid没有数据添加
    # print('df:\n', df)
    # df.drop(df.columns[[6, 7]], axis=1, inplace=True)
    address_db = parent_os + '/dbcalender.sqlite3'
    conn = sqlite3.connect(address_db)
    c = conn.cursor()
    list_update = []
    for lesson in df.values.tolist():
        # print(lesson)
        classtime = str(lesson[0])
        date = str(lesson[1])
        sql_search = "SELECT nid FROM app01_icalendar WHERE classtime='" + classtime + "' and date='" + date + "'"
        # print('sql_search:', sql_search)
        c.execute(sql_search)
        result_classname = c.fetchall()
        # print('result_classname:\n', result_classname)
        # print(len(result_classname))
        if len(result_classname) == 0:
            list_update.append(lesson)
    # print('list_update:', list_update)
    df_update = pd.DataFrame(list_update,
                             columns=['classtime', 'date', 'starttime', 'endtime', 'classname', 'created', 'dtstamp',
                                      'uid'])
    # print('\ndf_update:\n', df_update)
    print('\n数据库更新%s条数据\n' % (df_update.shape[0]))
    df_update.to_sql('app01_icalendar', con=conn, if_exists='append', index=False)
    conn.close()


def UID_Create():
    return random_str(20) + "&yzj"


def CreateTime():
    # 生成 CREATED
    global DONE_CreatedTime
    date = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")
    DONE_CreatedTime = date + "Z"


# 生成 UID
# global DONE_EventUID
# DONE_EventUID = random_str(20) + "&yzj"

# print("CreateTime")

def uniteSetting():
    #
    global DONE_ALARMUID
    DONE_ALARMUID = random_str(30) + "&yzj"
    #
    global DONE_UnitUID
    DONE_UnitUID = random_str(20) + "&yzj"


# print("uniteSetting")

def setClassTime():  # 打开class time文件
    data = []
    # print(conf_classTime_os)

    with open(conf_classTime_os, 'r', encoding='utf-8') as f:
        data = json.load(f)
    # print('\ndata_classTime:\n', data)
    global classTimeList
    classTimeList = data["classTime"]


def setClassInfo():  # 打开class info文件
    data = []
    with open(conf_classInfo_os, 'r', encoding='utf-8') as f:
        data = json.load(f)
    global classInfoList
    print('\ndata_classInfo:\n', data)
    classInfoList = data["classInfo"]


# print("setClassInfo:")

def setFirstWeekDate(firstWeekDate):
    global DONE_firstWeekDate
    DONE_firstWeekDate = time.strptime(firstWeekDate, '%Y%m%d')


# print("setFirstWeekDate:",DONE_firstWeekDate)

def setReminder(reminder):  # 设置提醒时间
    global DONE_reminder
    reminderList = ["-PT10M", "-PT30M", "-PT1H", "-PT2H", "-P1D"]
    if (reminder == "1"):
        DONE_reminder = reminderList[0]
    elif (reminder == "2"):
        DONE_reminder = reminderList[1]
    elif (reminder == "3"):
        DONE_reminder = reminderList[2]
    elif (reminder == "4"):
        DONE_reminder = reminderList[3]
    elif (reminder == "5"):
        DONE_reminder = reminderList[4]
    else:
        DONE_reminder = "NULL"


# print("setReminder",reminder)

def checkReminder(reminder):
    # TODO

    # print("checkReminder:",reminder)
    List = ["0", "1", "2", "3", "4", "5"]
    for num in List:
        if (reminder == num):
            return YES
    return NO


def checkFirstWeekDate(firstWeekDate):
    # 长度判断
    if (len(firstWeekDate) != 8):
        return NO;

    year = firstWeekDate[0:4]
    month = firstWeekDate[4:6]
    date = firstWeekDate[6:8]
    dateList = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    # 年份判断
    if (int(year) < 1970):
        return NO
    # 月份判断
    if (int(month) == 0 or int(month) > 12):
        return NO;
    # 日期判断
    if (int(date) > dateList[int(month) - 1]):
        return NO;

    # print("checkFirstWeekDate:",firstWeekDate)
    return YES


def basicSetting():
    info = "欢迎使用课程表生成工具。\n接下来你需要设置一些基础的信息方便生成数据\n"
    # print (info)

    info = "请设置第一周的星期一日期(如：20160905):\n"
    # firstWeekDate = input(info)

    firstWeekDate = '20210215'
    checkInput(checkFirstWeekDate, firstWeekDate)  # 检查初始周格式
    # checkInput(0, '20210208')

    info = "正在配置上课时间信息……\n"
    # print(info)
    # try :
    setClassTime()
    # print("配置上课时间信息完成。\n")
    # except :
    # 	sys_exit()

    info = "正在配置课堂信息……\n"
    # print(info)
    # try :
    setClassInfo()
    # print("配置课堂信息完成。\n")
    # except :
    # sys_exit()

    info = "正在配置提醒功能，请输入数字选择提醒时间\n【0】不提醒\n【1】上课前 10 分钟提醒\n【2】上课前 30 分钟提醒\n【3】上课前 1 小时提醒\n【4】上课前 2 小时提醒\n【5】上课前 1 天提醒\n"
    # reminder = input(info)
    reminder = '0'
    checkInput(checkReminder, reminder)


def checkInput(checkType, input):
    # checkFirstWeekDate = 0
    # checkReminder = 1
    if (checkType == checkFirstWeekDate):
        if (checkFirstWeekDate(input)):  # 判断输入格式
            info = "输入有误，请重新输入第一周的星期一日期(如：20160905):\n"
            firstWeekDate = input(info)
            checkInput(checkFirstWeekDate, firstWeekDate)
        else:
            setFirstWeekDate(input)  # 设置第一周并打印格式
    elif (checkType == checkReminder):
        if (checkReminder(input)):
            info = "输入有误，请重新输入\n【1】上课前 10 分钟提醒\n【2】上课前 30 分钟提醒\n【3】上课前 1 小时提醒\n【4】上课前 2 小时提醒\n【5】上课前 1 天提醒\n"
            reminder = input(info)
            checkInput(checkReminder, reminder)
        else:
            setReminder(input)

    else:
        print("程序出错了……")
    # end


def random_str(randomlength):
    str = ''
    chars = 'AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789'
    length = len(chars) - 1
    random = Random()
    for i in range(randomlength):
        str += chars[random.randint(0, length)]
    return str


importlib.reload(sys)
main()
