# -*- coding: utf-8 -*-
import xlrd
import xlwt
import sys
from xlwt import *    
from xlrd import open_workbook

from xlrd import open_workbook  
import sys  
  
#输出整个Excel文件的内容  
def print_workbook(wb):  
    for s in wb.sheets():  
        print("Sheet:", s.name)  
        for r in range(s.nrows):  
            strRow = ""  
            for c in s.row(r):
                #strRow += ("\t" + string(c.value))
                print(c.value)
    #print("ROW[" + r + "]:", strRow)  
  
#把一行转化为一个字符串  
def row_to_str(row):  
  strRow = ""  
  for c in row:  
      strRow += ("\t" + c.value)  
  return strRow;  
  
#打印diff结果报表  
def print_report(report):  
  for o in report:  
    if isinstance(o, list):  
      for i in o:  
        print("\t" + i)  
    else:  
      print (o)  
  
#diff两个Sheet  
def diff_sheet(sheet1, sheet2):  
  nr1 = sheet1.nrows  
  nr2 = sheet2.nrows  
  nr = max(nr1, nr2)  
  report = []  
  for r in range(nr):  
    row1 = None;  
    row2 = None;  
    if r<nr1:  
      row1 = sheet1.row(r)  
    if r<nr2:  
      row2 = sheet2.row(r)  
  
    diff = 0; # 0:equal, 1: not equal, 2: row2 is more, 3: row2 is less  
    if row1==None and row2!=None:  
      diff = 2  
      report.append("+ROW[" + str(r+1) + "]: " + row_to_str(row2))  
    if row1==None and row2==None:  
      diff = 0  
    if row1!=None and row2==None:  
      diff = 3  
      report.append("-ROW[" + str(r+1) + "]: " + row_to_str(row1))  
    if row1!=None and row2!=None:  
      # diff the two rows  
      reportRow = diff_row(row1, row2)  
      if len(reportRow)>0:  
        report.append("#ROW[" + str(r+1) + "]1: " + row_to_str(row1))  
        report.append("#ROW[" + str(r+1) + "]2: " + row_to_str(row2))  
        report.append(reportRow)  
  
  return report;  
  
#diff两行  
def diff_row(row1, row2):  
  nc1 = len(row1)  
  nc2 = len(row2)  
  nc = max(nc1, nc2)  
  report = []  
  for c in range(nc):  
    ce1 = None;  
    ce2 = None;  
    if c<nc1:  
      ce1 = row1[c]  
    if c<nc2:  
      ce2 = row2[c]  
      
    diff = 0; # 0:equal, 1: not equal, 2: row2 is more, 3: row2 is less  
    if ce1==None and ce2!=None:  
      diff = 2  
      report.append("+CELL[" + str(c+1) + ": " + ce2.value)  
    if ce1==None and ce2==None:  
      diff = 0  
    if ce1!=None and ce2==None:  
      diff = 3  
      report.append("-CELL[" + str(c+1) + ": " + ce1.value)  
    if ce1!=None and ce2!=None:  
      if ce1.value == ce2.value:  
        diff = 0  
      else:  
        diff = 1  
        report.append("#CELL[" + str(c+1) + "]1: " + ce1.value)  
        report.append("#CELL[" + str(c+1) + "]2: " + ce2.value)  
  
  return report  
  
  
'''if __name__=='__main__':  
  if len(sys.argv)<3:  
    exit()  
  
  file1 = sys.argv[1]  
  file2 = sys.argv[2]  
  
  wb1 = open_workbook(file1)  
  wb2 = open_workbook(file2)  
  
  #print_workbook(wb1)  
  #print_workbook(wb2)  
  
  #diff两个文件的第一个sheet  
  report = diff_sheet(wb1.sheet_by_index(0), wb2.sheet_by_index(0))  
  print file1 + "\n" + file2 + "\n#############################"  
  #打印diff结果  
  print_report(report)  
'''
#对比两个表格差异
#打开一个xls文件，读取数据
def open_excel(file= 'file.xls'):
    try:
        data = xlrd.open_workbook(file,encoding_override='utf-8')
        return data
    except Exception as e:
        print("文件打开错误")
        print(e)
''' 得到一个excel的sheet个数
    rb: 已经打开的excel对象
'''
def xl_sheet_num(rb):
    count = len(b.sheets()) #sheet数量
    return count


''' 获得一个excel所有的sheet名字
    rb: 已经打开的excel对象
'''
def xl_sheet_name(rb):
    count = len(rb.sheets())
    for sheet in rb.sheets():
        print(sheet.name)#sheet名称

'''获得表格中某个sheet某行的数据
    file：Excel文件路径
    colnameindex：行号
    by_index：sheet 号
'''
def excel_table_byindex(data,by_index=0,rowindex=0):
    #通过索引顺序获取一个表
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    print(nrows,ncols)
    if rowindex in range(1,nrows):
         #行列数据
        row = table.row_values(rowindex)
        print("row===",row)
        app = {}
        print("row_length==",len(row))
        return row
    else:
        return null
'''对比两个表格的差异
'''
def excel_table_compare(rb_hw,rb_hq):
    hw_sheet_num = xl_sheet_name(rb_hw)
    hq_sheet_num = xl_sheet_name(rb_hq)
    for i in range(hw_sheet_num):
        table = rb_hw.sheets()[i]
        nrows = table.nrows
        ncols = table.ncols
        for j in range(nrows):
            com_string(rb_hw,rb_hq,)
'''
    匹配关键字
    compare:需要对比的excel
    com_sheet：需要对比的sheet
    com_row_index:需要对比的行
    source:参考文件
    sour_sheet:参考sheet
    sour_row_index:参考文件行
    
'''
def com_string_row_col(compare,source,
               com_sheet=0,com_row_index=0,sour_sheet=0,
               sour_row_index=0):
    com_row = excel_table_byindex(compare,com_sheet,com_row_index)
    sour_row = excel_table_byindex(source,sour_sheet,sour_row_index)
    for i in range(2,len(com_row)-2):
        com_string = com_row[i];
        if (com_string==""):
            continue
        print("com_string==",com_string)
        if com_string in sour_row:
            return 1
    else:
        return 0
'''

'''
def com_string_row_sheet(rb_hw,rb_hq,hw_sheet,hw_row,hq_sheet):
    hq_table = data.sheets()[by_index]
    nrows = table.nrows
    for i in range(nrows):
        if (com_string_row_col(rb_hw,rb_hq,hw_sheet,hw_row,hq_sheet,i) ==1):
            return 1
    return 0

def main():
    hw_File = "C:\\Users\\zwx318792\\Desktop\\xls_test\\huawei.xls"
    hq_File = "C:\\Users\\zwx318792\\Desktop\\xls_test_change.xls"
    test_file = "/home/zsi1989u/zsl-github/zsl-excle/py_test.xlsx"

    #rb_hw = open_excel(hw_File)
    #rb_hq = open_excel(hq_File)
    rb_test = open_excel(test_file)
   
    print_workbook(rb_test)
    #print_workbook(rb_hw)
    #wb = Workbook()
    #list_row = excel_table_byindex(rb_hw,1108,1)
    #print(list)
    #isinclude = com_string(rb_hw,rb_hw,5,87,3,67)
    #print(isinclude)


if __name__=="__main__":
    main()
