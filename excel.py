# _*_ coding:utf-8 _*_
import xlrd
import xlwt
import pandas as pd
data = xlrd.open_workbook('data.xlsx')
table = data.sheets()[0]

data_new=xlwt.Workbook()
table_new = data_new.add_sheet('wangmiao')
table_list=[]
nrows=table.nrows

for i in range(1,nrows):
    col_o=[]
    col_p=[]
    col_q=[]
    if i==10:
        table_list.append(table.row_values(i))
        continue

    col_o= table.cell_value(i,14)
    col_o= col_o.encode('unicode-escape').decode('string_escape')
    col_o=col_o.split(',')

    col_p = table.cell_value(i,15)
    col_p = col_p.encode('unicode-escape').decode('string_escape')
    col_p = col_p.split(',')

    col_q = table.cell_value(i,16)
    col_q = col_q.encode('unicode-escape').decode('string_escape')
    col_q = col_q.split(',')


    
    for j in range(len(col_o)):
        row_i=table.row_values(i)
        row_i[14]=col_o[j]
        row_i[15]=col_p[j]
        row_i[16]=col_q[j]
        table_list.append(row_i)

output=open('neweeeew.xlsx','w')
output.write('  \tbaitID\toeID\tbaitChr\tbaitstart\tbaitend\tbaitName\tc10s1\tc10s2\tc10s3\tc10s4\tmin_weighted_padj\tdaltaAsinhScore\tregionIDs\tlog2FoldChanges\tweighted\tOEranges\n')

for a in range(len(table_list)):
    for b in range(len(table_list[a])):
        output.write(str(table_list[a][b]))
        output.write('\t')
    output.write('\n')
output.close()



