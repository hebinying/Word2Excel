#coding=utf-8
import xlrd,xlwt
import readWord

C=["case_id","checkname","method","url","params","checkpoint"]


class writeExcel():
    def __init__(self,datas):
        self.datas=datas
        self.wb=xlwt.Workbook()
        self.sheet=self.wb.add_sheet((datas["wordname"]).decode("gbk"))
        #分解数据
        self.filename=datas["wordname"]
        self.filepath=datas["wordpath"]
        self.tables=datas["content"]
        self.tablenumber=datas["tablenumber"]
    #分解content数据并保存到excel表中
    def write_tabledata(self):
        num=0
        dictname={}
        for i in range(len(C)):
            self.sheet.write(num,i,C[i])

        for i in range(len(self.tables)):
            num+=1
            dictname = {}
            if len(self.tables[i])!=0:
                dictname["case_id"]="case_"+str(num)
                dictname["checkname"]=(self.tables[i]["testcasename"])
                dictname["method"]=self.tables[i]["method"]
                dictname["url"]=self.tables[i]["url"]
                dictname["params"]=(self.tables[i]['params'])
                dictname["checkpoint"]=self.tables[i]['checkpoint']
                #dictname["sample"]=(self.tables[i]['sample'])
            self.write_excel(num,dictname)
        self.wb.save(self.filename+'.xlsx')

    def write_excel(self,rownum,data):
        for i in range(len(C)):
            # print data[C[i]]
            self.sheet.write(rownum,i,data[C[i]])






if __name__=="__main__":
    a = readWord.WordUtil("m2c.scm.api-1.3.0.docx")
    data=a.get_tablesdata()
    b=writeExcel(data)
    b.write_tabledata()
