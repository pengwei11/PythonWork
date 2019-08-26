# -*- coding:utf-8 -*-
'''
Auth：彭威
date:2018年09月28日17时
title:将数据写入excel
缺陷：
未做分模块写入
'''
import os
import xlwt
from xmindparser import xmind_to_json
import json,jsonpath
import win32ui





# 数据输入
print("请输入xmind路径（不能有中文路径）")
dlg = win32ui.CreateFileDialog(1) # 1表示打开文件对话框
dlg.SetOFNInitialDir('E:/Python') # 设置打开文件对话框中的初始显示目录
dlg.DoModal()
filename = dlg.GetPathName() # 获取选择的文件名称
print(filename)
print("请输入测试类型")
TestType = input()
print("请输入测试模块")
TestModule = input()
print("请输入测试子项目")
TestSubItem = input()
print("请输入测试用例编号")
TestCaseID = input()
print("请输入需求编号")
DemandNumber = input()
print("请输入测试标题")
TestTitle = input()
print("请输入重要级别")
ImportantLevel = input()
print("请输入预置条件")
Precondition = input()
print("请输入执行步骤")
ExecutionSteps = input()
print("请输入输入数据")
InputData = input()
print("请输入预期输出")
ExpectedOutput = input()
print("请输入兼容硬件")
CompatibleHardware = input()


# 读取xmind并转换json
def xmind_Out():
    out_file = xmind_to_json(filename)
    return out_file

# 将json文件转换为可读文件
def json_read():
    test = open(os.path.realpath(xmind_Out()),'r') # 打开json文件
    load_dic = json.loads(test.read())  # 将json格式转换为可读文件
    title_list = jsonpath.jsonpath(load_dic,"$..title") # 读取load_dic中key为title的value
    title_list1 = []
    for title in title_list:
        title_list1.append(title)
    return title_list1

# 创建excel
def wrilt_excel():
    # 创建excel
    workbook = xlwt.Workbook(encoding='UTF-8')
    # 创建工作表
    sheet1 = workbook.add_sheet('工作表请自行创建')
    # 第一，二行数据
    row_one = ['用例总数','','通过','','不通过','','未测','','新增用例数','','','','','','','','','','','','','','','','','']
    row_two = ['测试类型', '测试模块', '测试子项目', '测试用例编号', '需求编号', '测试标题', '重要级别', '预置条件', '执行步骤', '输入数据', '预期输出', '兼容硬件', '测试结果', 'bug编号', '实际结果', '备注', '是否是新增用例', '用例增加日期', '用例编写人', '用例增加版本', '是否自动化', '自动化编号', '是否稳定性测试', '稳定性测试频度', '功能分类', '测试角色']
    title = json_read()

    for i in range(len(row_one)):
            sheet1.write(0, i, row_one[i],xlwt.easyxf('font:height 200, name Arial_Unicode_MS, colour_index black, bold on;align: horiz center;'))
            sheet1.write(1,i,row_two[i],xlwt.easyxf('font:height 200, name Arial_Unicode_MS, colour_index black, bold on;align: horiz center;'))

    for j in range(len(title)):
        sheet1.write(j + 2, 0,TestType)
        sheet1.write(j + 2, 1,TestModule)
        sheet1.write(j + 2, 2, TestSubItem)
        sheet1.write(j + 2, 3, TestCaseID)
        sheet1.write(j + 2, 4, DemandNumber)
        sheet1.write(j + 2, 5, TestTitle)
        sheet1.write(j + 2, 6, title[j])
        sheet1.write(j + 2, 7, ImportantLevel)
        sheet1.write(j + 2, 8, Precondition)
        sheet1.write(j + 2, 9, ExecutionSteps)
        sheet1.write(j + 2, 10, InputData)
        sheet1.write(j + 2, 11, ExpectedOutput)
        sheet1.write(j + 2, 12, CompatibleHardware)




    workbook.save('C:/Users/Administrator/Desktop/demo.xlsx')

if __name__=='__main__':
    xmind_Out()
    json_read()
    wrilt_excel()
    print("创建demo.xlsx文件成功")