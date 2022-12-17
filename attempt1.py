#输入一个字符串，判断里面是否含有相同字母，输出bool

from ast import Num
from pickle import FALSE
import string


def show():
    str0="abcab"                 #不能用str来命名
    length=len(str0)
    comp={string:bool}
    for i in range(0,length):
        num1 = str0[i]
        ii=str(i)                #int 转 string
        for j in range(i+1,length):
            num2 = str0[j]
            jj=str(j)            #int 转 string
            ij=ii+'*'+jj
            if num1 == num2:
                comp[ij]=True
            #else:
                #comp[ij]=False
    if len(comp) == 1:
        print ('false')
    else:
        print ('true')
        #for comp_result in comp.keys():     #首个占用？
            #print (comp_result+"重复")
if __name__=="__main__":
     show()