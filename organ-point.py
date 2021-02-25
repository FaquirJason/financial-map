import os, re
import csv
import pandas as pd
import numpy as np

# import xlwt

def read_csv(pas,file): 
    ori_rows = pd.read_csv(file)
    ori_rows = ori_rows.values.tolist()

    uid = []
    content = []
    all_opinions = []

    for row in ori_rows:
        uid.append(row[0])
        content.append(row[1])
        opinions = eval(row[2])

        all_opinions.append(opinions)

    return uid, all_opinions 
    

if __name__ == '__main__':
    uid, opinions = read_csv(".","1111.csv")
    # print(contents)
    i = 0
    # for content in contents:

    # workbook = xlwt.Workbook(encoding = 'utf-8')
    # worksheet = workbook.add_sheet('sheet')

    i = 0
    length = len(opinions)
    # count = 0

    with open("Triples.csv",'w') as f:
        csv_write = csv.writer(f)
        while i < length:
            dic = opinions[i]
            keys = list(dic.keys())
            # print(keys) 
            j = 0
            while j < len(keys):
                # len(keys)
                k = 0
                # print(dic[keys[j]])
                while k < len(dic[keys[j]]):
                    # len(dic[keys[j]])
                    csv_write.writerow(["V","organization",str(keys[j]),str(keys[j])])
                    # worksheet.write(3*count,0,"V")
                    # worksheet.write(3*count,1,"organization")
                    # worksheet.write(3*count,2,keys[j])
                    # worksheet.write(3*count,3,keys[j])
                    csv_write.writerow(["V","opinion",str(uid[i]+str(j)+str(k)),str(dic[keys[j]][k])])
                    # worksheet.write(3*count+1,0,"V")
                    # worksheet.write(3*count+1,1,"opinion")
                    # worksheet.write(3*count+1,2,uid[i]+str(j)+str(k))
                    # worksheet.write(3*count+1,3,str(dic[keys[j]][k]))
                    csv_write.writerow(["E","发布",str(keys[j]),str(uid[i]+str(j)+str(k))])
                    # worksheet.write(3*count+2,0,"E")
                    # worksheet.write(3*count+2,1,"发布")
                    # worksheet.write(3*count+2,2,keys[j])
                    # worksheet.write(3*count+2,3,uid[i]+str(j)+str(k))
                    k += 1
                    # count += 1
                j += 1
            i += 1


