# -*- coding: utf-8 -*- 
# required: xlrd, xlwt, xlutils
# xlrd: read xls
# xlwt: write xls
# xlutils: modified xls , need xlrd & xlwt
#
__author__ = 'yelord'

import os
import xlrd, xlwt
from xlutils.copy import copy
from xlwt import easyxf

# xlsFile = 'excelFile.xls'

# 获取一个工作表
# data = xlrd.open_workbook(xlsFile)
# table = data.sheets()[0]

# 获取整行和整列的值（数组）
# table.row_values(ai_row)
# table.col_values(ai_col)

# 获取行数和列数
# nrows = table.nrows
# ncols = table.ncols

# 单元格
# cell_A1 = table.cell(0,0).value
# cell_C4 = table.cell(2,3).value
#
# 使用行列索引
# cell_A1 = table.row(0)[0].value
# cell_A2 = table.col(1)[0].value

# 类型 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
# ctype = 1 value = '单元格的值'
#
# xf = 0 # 扩展的格式化
# table.put_cell(row, col, ctype, value, xf)
# table.cell(0,0)  #单元格的值'
# table.cell(0,0).value #单元格的值'


def ReadXls(filename):
    list = []
    try:
        data = xlrd.open_workbook(filename)
    except:
        return list
    sheet = data.sheets()[0]
    ai_rows = sheet.nrows
    ai_cols = sheet.ncols

    if ai_cols == 0 or ai_rows == 0:
        return list

    for i in xrange(0,ai_rows ):
        type = sheet.cell(i,0).ctype
        if type == 1 :  #string
            list.append(sheet.cell(i,0).value)
        if type == 2 :  #number
            list.append(unicode(long(sheet.cell(i,0).value)))

    return list

def ReadText(filename):
    from sku import as_unicode
    list = []
    try:
        list = [as_unicode(barcode.strip()) for barcode in open(filename).readlines() if barcode.strip()]
    except:
        pass
    return list

def ReadData(filename):
    s = os.path.splitext(filename)[-1].lower()

    if s == '.txt':
        return ReadText(filename)
    elif s == '.xls':
        return ReadXls(filename)
    else:
        return []

def WriteXls_demo(filename):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Ye-test sheet')
    worksheet.write(0, 0, 5) # Outputs 5
    worksheet.write(0, 1, 2) # Outputs 2
    worksheet.write(1, 0, xlwt.Formula('A1*B1')) # Should output "10" (A1[5] * A2[2])
    worksheet.write(1, 1, xlwt.Formula('SUM(A1,B1)')) # Should output "7" (A1[5] + A2[2])
    workbook.save(filename)

def ModifyXls(filename):
    # formatting_info=True: 保存之前数据的格式
    # on_demand=True: 节省内存
    rb_workbook = xlrd.open_workbook(filename, formatting_info=True, on_demand=True)
    # copy the xlrd.Book object into an xlwt.Workbook object
    wb_workbook = copy(rb_workbook)
    sheet = wb_workbook.get_sheet(0)
    sheet.write(3,0,100)
    sheet.write(3,2,'ok!')
    sheet.write(3,1,u'好')
    print sheet.nrows
    wb_workbook.save(filename)

def AddContentFromXls(basefile, fromfile, show_tips):

    try:
        rb = xlrd.open_workbook(basefile,formatting_info=True)
    except:
        show_tips(u' ！！！打开基础文件失败！！！')
        return
    sheet = rb.sheets()[0]
    ai_rows = sheet.nrows

    # if ai_rows == 0:
    #     return

    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    baselist = ReadXls(basefile)

    try:
        f_rb = xlrd.open_workbook(fromfile, formatting_info=True)
    except:
        show_tips(u' ！！！打开追加文件失败！！！')
        return

    # xlwt.easyxf('font: color-index red, bold on')

    f_sheet = f_rb.sheets()[0]
    f_rows = f_sheet.nrows
    cnt = 0
    for i in xrange(0,f_rows ):
        type = f_sheet.cell(i,0).ctype
        barcode =''
        # string
        if type == 1:
            barcode = f_sheet.cell(i,0).value
        elif type == 2:  #number
            barcode = unicode(long(f_sheet.cell(i,0).value))

        if barcode:
            try:
                idx = baselist.index(barcode)
            except:
                idx = -1
            # 条码不存在
            if idx == -1:
                baselist.append(barcode)
                for j  in xrange(0,f_sheet.ncols):
                    if j==0:
                        w_sheet.write(ai_rows,j, barcode)
                    else:
                        w_sheet.write(ai_rows,j, f_sheet.cell(i,j).value)

                ai_rows+=1
                cnt+=1
                tips = u'第 %s 行条码已加上(%s),ok！'%(i,barcode)
            else:
                tips = u'第 %s 行条码已存在(%s),pass'%(i,barcode)

            show_tips(tips)
    wb.save(basefile)
    tips = u'== 完成！文件(%s)已增加 %s 条新记录！从 %s 到 %s 行。'%(basefile,cnt,ai_rows - cnt + 1, ai_rows )
    show_tips(tips)

def show(tips):
    print tips


if __name__ == '__main__':
    # print ReadXls(u'/Users/yelord/window/tmp/sku/demo数据和图片/demo数据和图片/demo.xls')
    # print ReadText(u'/Users/yelord/PycharmProjects/sku_info/test.txt')
    #
    # print ReadData(u'/Users/yelord/PycharmProjects/sku_info/test.txt')
    #
    # WriteXls_demo('xlwt_test.xls')
    # print 'xlwt_test.xls created!'
    #
    # ModifyXls('xlwt_test.xls')
    AddContentFromXls(u'/Users/yelord/window/tmp/sku/demo数据和图片/demo数据和图片/demo.xls',
                      u'/Users/yelord/window/tmp/sku/demo数据和图片/demo数据和图片/demo2.xls',
                      show)
