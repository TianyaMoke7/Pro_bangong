# 导图纸调用函数

import os
from win32com.shell import shell,shellcon
 
# debug = False
def fuzhi(filename1,filename2):#filename1是原路径，filename2是要保存的路径
    print('开始下载', filename1,'下载中...')
    # if not debug:
    res = shell.SHFileOperation((0, shellcon.FO_COPY, filename1, filename2,
                                  shellcon.FOF_NOCONFIRMATION | shellcon.FOF_NOERRORUI | shellcon.FOF_SILENT,
                                 None, None))

if __name__ == '__main__':            # 当函数fuzhi（）被其他模块调用时，不执行以下内容
    path_yuan = r'D:\调试文件夹\old'
    path_xian = 'D:\\调试文件夹\\new'
    fuzhi(path_yuan,path_xian)










# 字体背景颜色对照表：
# https://blog.csdn.net/zkw_1998/article/details/103930052?ops_request_misc=&request_id=&biz_id=102&utm_term=python%20EXCEL%E8%83%8C%E6%99%AF%E9%A2%9C%E8%89%B2&utm_medium=distribute.pc_search_result.none-task-blog-2~all~sobaiduweb~default-0-103930052.142^v9^pc_search_result_control_group,157^v4^control&spm=1018.2226.3001.4187


################################################# 填充颜色实例 #########################################
# 例程代码，不参与本项目
# 地址：https://blog.csdn.net/bigfishfish/article/details/123247362?ops_request_misc=&request_id=&biz_id=102&utm_term=python%20openpyxl%20%E9%A2%9C%E8%89%B2&utm_medium=distribute.pc_search_result.none-task-blog-2~all~sobaiduweb~default-0-123247362.142^v9^pc_search_result_control_group,157^v4^control&spm=1018.2226.3001.4187

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
#             ws.cell(row+1,col).fill=PatternFill(patternType=pattern_Type[row-1],start_color=Color(index=2))
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
################################################# 填充颜色实例 #########################################