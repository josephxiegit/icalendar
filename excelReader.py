# coding: utf-8
#!/usr/bin/python

import sys
import xlrd 
import importlib
import os
from datetime import datetime

__author__ = 'joseph'

xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True
# 指定信息在 xls 表格内的列数
_colOfClassName = 0
_colOfStartWeek = 1
_colOfEndWeek = 2
_colOfWeekday = 3
_colOfClassTime = 4
_colOfClassroom = 5
_colOfDate = 6
_colOfStartTime = 7
_colOfEndTime = 8


def main():
	# 读取 excel 文件
	parent = os.path.dirname(os.path.realpath(__file__))
	workbook = xlrd.open_workbook(parent + "/calendar_excel_input.xlsm")
	table = workbook.sheets()[0]
	
	# 基础信息
	numOfRow = table.nrows  #获取行数,即课程数
	numOfCol = table.ncols  #获取列数,即信息量
	headStr = '{\n"classInfo":[\n'
	tailStr = ']\n}'
	classInfoStr = ''
	classInfoArray = []
	# 信息列表
	# lengthOfList = numOfRow-1
	# print('lengthOfList', lengthOfList)
	classNameList = []
	startWeekList = []
	endWeekList = []
	weekdayList = []
	classTimeList = []
	classroomList = []
	dateList = []
	startTimeList = []
	endTimeList = []


	# 开始操作

	def handleTime(str, flag):
		# 将时间处理成标准格式
		# print(str)
		if flag == 0: # 处理开始时间
			str = str.replace('点', '')
			str = str.replace('半', '30')
			if len(str) == 1:
				str = "0" + str + "00"
			if len(str) == 2:
				str = str + "00"
			if len(str) == 3:
				str = "0" + str
			# print(str)
			return str
		else:			# 处理结束时间	
			str = int(str) + 200
			# print(str)
			return repr(str)

	# 将信息加载到列表
	i = 1
	while i < numOfRow:
		index = i-1

		classNameList.append(((table.cell(i, _colOfClassName).value)))
		startWeekList.append(str(int((table.cell(i, _colOfStartWeek).value))))
		endWeekList.append(str(int((table.cell(i, _colOfEndWeek).value))))
		weekdayList.append(str(int((table.cell(i, _colOfWeekday).value))))
		classTimeList.append(str(int((table.cell(i, _colOfClassTime).value))))
		classroomList.append(str(((table.cell(i, _colOfClassroom).value))))

		startTimeNew = handleTime(str((table.cell(i, _colOfStartTime).value)), 0)
		startTimeList.append(startTimeNew)

		endTimeNew = handleTime(startTimeNew, 1)
		endTimeList.append(endTimeNew)
		
		dateList.append(str(table.cell(i, _colOfDate).value))
		
		i += 1
	i = 0
	itemHeadStr = '{\n'
	itemTailStr = '\n}'

	classInfoStr += headStr
	classInfoStr_len = len(classNameList)
	for className in classNameList:
		itemClassInfoStr = ""
		itemClassInfoStr  = itemHeadStr + '"className":"' + className + '",'
		itemClassInfoStr += '"week":{\n"startWeek":' + startWeekList[i] + ',\n'
		itemClassInfoStr += '"endWeek":' + endWeekList[i] + '\n},\n'
		itemClassInfoStr += '"weekday":' + weekdayList[i] + ',\n'
		itemClassInfoStr += '"classTime":' + classTimeList[i] + ',\n'
		itemClassInfoStr += '"classroom":"' + classroomList[i] + '",\n'
		itemClassInfoStr += '"startTime":"' + startTimeList[i] + '",\n'
		itemClassInfoStr += '"endTime":"' + endTimeList[i] + '",\n'

		# 修正日期格式
		# print(dateList[i])
		date_new = int(float(dateList[i]))
		date_new = xlrd.xldate_as_datetime(date_new, workbook.datemode)
		date_new = datetime.strftime(date_new, "%Y%m%d")
		print(date_new)

		itemClassInfoStr += '"date":' + date_new + '\n'

		itemClassInfoStr += itemTailStr
		classInfoStr += itemClassInfoStr
		if i!=len(classNameList)-1 :
			classInfoStr += ","
		i += 1

	classInfoStr += tailStr
	print('classInfoStr:', classInfoStr)
	print("\n合计共 %s 条数据\n" % (classInfoStr_len))

	parent = os.path.dirname(os.path.realpath(__file__))
	conf_classInfo = parent + "/conf_classInfo.json"
	
	with open(conf_classInfo,'w') as f:
		f.write(classInfoStr)
		f.close()
	print("saved as path %s\n" % (conf_classInfo))

importlib.reload(sys)

main()