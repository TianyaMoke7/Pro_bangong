# # 爬取bilibili游戏排行榜名称和链接
# import requests
# from bs4 import BeautifulSoup

# 各种小尝试

# import xlrd
# from xlutils.copy import copy
# import xlwt
# rb = xlrd.open_workbook('linux.xlsx')
# wb = copy(rb)
# ws = wb.get_sheet(0)
# for i in range(2,159):
#     ws.write(i,1,"=HYPERLINK(\"#sheet2!a{}\")\r".format(i))
# wb.save('linux.xls')


# import re
# line="this hdr-biz model server"
# pattern=r"hdr-biz"
# m = re.search(pattern, line)
# print(m) 


# import pandas as pd
# def find_row(num_value,file_name):
# # Returns the row number based on the value of the specified cell
#     demo_df = pd.read_excel(file_name)
#     for indexs in demo_df.index:
#         for i in range(len(demo_df.loc[indexs].values)):
#             if (str(demo_df.loc[indexs].values[i]) == num_value):
#                 row = str(indexs+2).rstrip('L')
#                 return row
# row_num = find_row('103.03','test.xlsx')
# print(row_num)

# index()用法
# A = [123, 'xyz', 'zara', 'abc']
# print(A.index('zara')) 


#查找操作
# ef find_files(file):
#     find_file = input("请输入需要模糊查找的文件名：")
#     end_file = input("文件类型为(如果没有文件后缀要求或查找要素有文件夹直接回车即可)：")
#     for f in file:
#         if find_file in f and f.endswith(end_file):
#             print(f"找到文件名中包括{find_file},并且文件类型是{end_file}的完整文件名字有：{f}")


# # 模糊查找
# import os
# # 路径（鼠标右键查看文件属性）
# path = 'D:\网页下载'
# files = os.listdir(path)
# # 查找文件名字含有fish且以.png后缀的文件
# for f in files:
#     if '研发' in f and f.endswith('.pptx'):
#         print('Found it!' + f)


################################################# 填充颜色实例 #########################################
# from openpyxl import Workbook
# from openpyxl.styles import *
 
# save_file_path="xxxxxxxxxx.xlsx"
# wb = Workbook()
# ws = wb.active
# pattern_Type =['darkDown', 'darkUp', 'lightDown', 'darkGrid', 'lightVertical',
#                'solid', 'gray0625', 'darkHorizontal', 'lightGrid', 'lightTrellis',
#                'mediumGray', 'gray125', 'darkGray', 'lightGray', 'lightUp',
#                'lightHorizontal', 'darkTrellis', 'darkVertical']
# #设置列宽和行高
# ws.column_dimensions['a'].width = 15
# ws.column_dimensions['b'].width = 25
# ws.column_dimensions['c'].width = 25
# ws.column_dimensions['d'].width = 25
# for i in range(1,20):
#     ws.row_dimensions[i].height=25
# #列标题
# ws.cell(1,2).value="fgColor或start_color"
# ws.cell(1,3).value="bgColor或end_color"
# ws.cell(1,4).value="fgColor,bgColor"
# #填充单元格颜色
# for col in range(1,5):     # 列
#     for row in range(1, 19):   #行
#         if col==1:
#             ws.cell(row+1,col).value=pattern_Type[row-1]
#         elif col==2:
#             ws.cell(row+1,col).fill=PatternFill(patternType=pattern_Type[row-1],start_color=Color(index=42))
#         elif col==3:
#             ws.cell(row+1,col).fill=PatternFill(patternType=pattern_Type[row-1],start_color=Color(index=0), end_color=Color(index=4))
#         else:
#             ws.cell(row+1,col).fill=PatternFill(patternType=pattern_Type[row-1],start_color=Color(index=2), end_color=Color(index=3))
 
# wb.save(save_file_path) 
# patternType是指单元格填充底纹纹路  样式，
# fgColor前景色就是指花纹纹路的颜色是所设定值，
# bgColor背景色是指花纹纹路为默认黑色，单元格背景颜色为所设定值
# 特殊的是'solid'样式，也就是纯色填充样式，fgColor+bgColor时，由于fgColor是完全纯色，也就是一整块纹路，把bgColor给完全覆盖了，
# 所以效果跟只有fgColor一样，即只看到前景（花纹）红色，看不到背景（单元格）绿色。


#############################################   pandas表格查重   ###################################### 未实现
#1、导入包
# import pandas as pd
# #2、找到文件
# s = pd.read_excel(r"C:\Users\wu_jianguo\Desktop\Small Tools\实验图纸号.xlsx")

# #3、用drop_duplicates方法
# # s.drop_duplicates(subset="Name",inplace=True)
# #查重复，返回True or False
# dupe = s.duplicated(subset="Name")
# #过滤出重复的元素
# dupe = dupe[dupe==True]
# #根据ID,找到所在行
# print(s.iloc[dupe.index])

###############################################   阵列尝试   #######################################
# r1=[1,2,3,4,5]
# r2=[10,20,30,40,50]
# r3=[100,200,300,400,500]
# arr=[r1,r2,r3]
# for row in arr:
#     print(row)
#     for one in row:
#         print(one)

# s=(sum(row) for row in arr) #创建一个逐行统计的生成器
# i=0
# while (i<len(arr)):
# 	print ("Line %d, sum=%d" %(i , next(s))) 
# 	i=i+1



# from openpyxl import load_workbook
# # from openpyxl import Workbook
# from openpyxl.styles import *
# a = 1
# r1=[1,2,3,4,5]
# r2=[10,20,30,40,50]
# r3=[100,200,300,400,500]
# arr=[r1,r2,r3]
# for row in arr:
#     print(row)
#     for one in row:
#         print(one)
# wb = load_workbook(r'C:\Users\wu_jianguo\Desktop\Small Tools\新建 XLSX 工作表.xlsx')
# sheet1 = wb.worksheets[0]
# i = 1
# j = 1
# sheet1.cell(4+i,99+j,1234)
# sheet1.cell(4+i,99+2,1234).fill=PatternFill(patternType= 'solid',start_color=Color(index=5))
# sheet1.cell(4+i,99+3,1234).fill=PatternFill(patternType= 'solid',start_color=Color(index=13))
# sheet1.cell(4+i,99+4,1234).fill=PatternFill(patternType= 'solid',start_color=Color(index=34))
# sheet1.cell(4+i,99+5,1234).fill=PatternFill(patternType= 'solid',start_color=Color(index=43))


# wb.save(r'C:\Users\wu_jianguo\Desktop\Small Tools\新建文件夹\生产计划请购大表.xlsx')

# print('完成')
# print('完成')



for x in range(1, 10):
    print(x)    
    if x == 4:        
        break
print (x)