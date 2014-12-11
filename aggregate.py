#!/usr/bin/env python
#coding=utf-8

## author:chenke
## contact:chenke1818@gmail.com


import sys
import xlrd
import re
reload(sys)
sys.setdefaultencoding( "utf-8" )
def calculate(path1,keyindex1,valindex1,path2,keyindex2,valindex2):
    s1map={}
    s2map={}
    data1 = xlrd.open_workbook(path1)
    table1 = data1.sheets()[0]
    keys1 = table1.col_values(int(keyindex1)-1)
    vals1 = table1.col_values(int(valindex1)-1)
    i = 0
    for k in keys1:
        if k and k != '':
          #print k
          m = re.match(r'([0-9]*)([^\d]+)', k)
          kk = m.groups()[1]
          if s1map.get(kk,0.0) > 0.0:
              s1map[kk]  =  s1map[kk] + vals1[i]
          else:
              s1map[kk]  =  vals1[i]
        i = i + 1


    data2 = xlrd.open_workbook(path2)
    table2 = data2.sheets()[0]
    keys2 = table2.col_values(int(keyindex2)-1)
    vals2 = table2.col_values(int(valindex2)-1)
    i = 0
    for k in keys2:
        if k and k != '' and '汇总' in k:
          kk = k.replace('汇总','').strip()
          if s2map.get(kk,0.0) > 0.0:
              s2map[kk]  =  s2map[kk] + vals2[i]
          else:
              s2map[kk]  =  vals2[i]
        i = i + 1

    for k,v in s1map.items():
       print k,v,s2map.get(k)

if __name__ == '__main__':
      if len(sys.argv)!=7:
           print "参数格式为：python aggregate.py excel文件1路径 key的列号 值的列号 excel文件2路径 key的列号 值的列号"
      else:
           calculate(sys.argv[1],sys.argv[2],sys.argv[3],sys.argv[4],sys.argv[5],sys.argv[6])
