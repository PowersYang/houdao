# coding=utf-8

import xlrd
import requests
import json


if __name__ == '__main__':
    excelFile = '/Users/PowerYang/Desktop/yinchuan/厚道店面统计.xlsx'
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    f = open('/Users/PowerYang/Desktop/yinchuan/address.json', 'w+')

    for rowNum in range(table.nrows):
        if (rowNum not in [0, 1, 2, 3]):    
            rowVale = table.row_values(rowNum)
            mall = rowVale[1]
            address = '银川市' + rowVale[2]
            
            url = 'https://restapi.amap.com/v3/geocode/geo?address={address}&output=JSON&key=b270aab18c92b82a903ca0f4b3e50d04'.format(address=address + mall)

            res = requests.get(url)
            j = json.loads(res.text)
            if (j['status'] == '1'):
                location = j['geocodes'][0]['location']
                s = {"mall":mall, "address":address + mall, "location": location}
                
                f.write(str(s))
                f.write('\n')
            else:
                print(j)
                print(mall, address)                    
            print("---------------")
    f.close()
