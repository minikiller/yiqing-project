#coding=utf-8

import configparser
import re
import openpyxl				#导入openpyxl


class Files:
    def __init__(self, data, filename,sheet,input):
        self.data = data
        self.files = filename
        self.sheet=sheet
        self.inputData=input

dicts={}

# 解析config文件
def getCofing():
    config = configparser.ConfigParser()
    config.read('./config.ini')
    # print(config.sections())
    index=config['DEFAULT']['key']  # key=A1,A2
    keys=index.split(",")   #['A1','A2']
    
    print(keys)
    for key in keys:
        # print("--"+key)
        values=config[key]['key'] 
        # print("values ",values)
        strs=values.split(",")
        # 保存搜索的匹配串
        data=[] 
        #excel写入的坐标
        # pos=[]
        #excel file name
        filename=""
        for str in strs:  # D1=驻长?所高校 1,2
            k=config[key][str].split(" ")
            data.append(config[key][str]) # 驻长?所高校 
            # pos.append(k[1]) # 1,2
            filename=config[key]["filename"]
            sheet=config[key]["sheet"]
            input=config[key]["input"]
         
        dicts[key]=Files(data, filename,sheet,input)
        # print()
    print(dicts)

def parseData(index):
    result={}
    target=dicts[index]
    values=target.data
    search=target.inputData
    # print(data1[0])
    for value in values:
        mystr=value.split(" ") 
        target=mystr[0].replace("?","(.+?)")
        n = re.findall(target, search)
        result[mystr[0]]=n[0] 
    # print(result)
    for key, value in result.items():
        print(key, ': ', value)
    return result

def writeExcel(key,result):
    sheetname =int(dicts[key].sheet)
    # frame = pd.read_excel("./data.xlsx",int(sheet))
    filename='./'+dicts[key].files
    wb = openpyxl.load_workbook( filename)	
    # print(wb.get_sheet_names())
    # # 根据sheet名字获得sheet
    # a_sheet = wb.get_sheet_by_name(dicts[key].sheet)
    # # 获得sheet名
    # print(a_sheet.title)
    # # 获得当前正在显示的sheet, 也可以用wb.get_active_sheet()
    # sheet = wb.active
    # n = 0
    sheets = wb.sheetnames
    sheet = wb[sheets[sheetname]]
    # print(frame.info())
    data = dicts[key].data
    for value in data:
        list=value.split(" ") 
        key=list[0]
        target=list[1] 
        res=int(result.get(key))
        print("write data {} to {}".format(res,target))
        sheet[target]=res
        # frame.iloc[i,j] = res
    wb.save(filename)
    # frame.to_excel("./data1.xlsx")
    # df.to_excel("./eat1.xlsx")
    # for key, value in data.items():
    #     print(key, ': ', value)
    # frame.iloc[1,3] = '老王'


def getInput():
    key=input("请选择关键点信息，例如（A1）:")
    # input_str=input("请输入信息:")
    # print("data type is ", input_str )
    # print("data ",key ," value is ",input_str)
    return key
    # return key,input_str


if __name__ == "__main__":
    # getInput()
    # input="驻长10所高校，共报告295例（教职工23例、学生272例），其中：确诊6例（学生6例），校内隔离4855人、居家隔离669人、集中隔离8441人；密接7762人、次密接4352人。"
    # key="A1"
    key=getInput()
    getCofing()
    result=parseData(key)
    writeExcel(key,result)
