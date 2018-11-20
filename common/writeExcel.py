#coding=utf-8
import xlrd,xlwt
import readWord

C=["case_id","checkname","method","url","params","checkpoint"]


class writeExcel():
    def __init__(self,datas):
        self.datas=datas
        self.wb=xlwt.Workbook()
        self.sheet=self.wb.add_sheet((datas["wordname"]).decode("gbk"))
        #############
        self.col0=self.sheet.col(0)
        self.col1 = self.sheet.col(1)
        self.col2 = self.sheet.col(2)
        self.col3 = self.sheet.col(3)
        self.col4 = self.sheet.col(4)
        self.col5 = self.sheet.col(5)
        self.col0.width=230*10
        self.col1.width = 300 * 20
        self.col2.width = 230 * 10
        self.col3.width = 300 * 25
        self.col4.width = 300 * 25
        self.col5.width = 300 * 30
        ############
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
        ###############增加自动换行格式
        alignment = xlwt.Alignment()
        # alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        alignment.vert = xlwt.Alignment.VERT_CENTER
        style = xlwt.XFStyle()
        style.alignment = alignment
        ###########

        for i in range(len(C)):
            # print data[C[i]]
            self.sheet.write(rownum,i,data[C[i]],style)






if __name__=="__main__":
    a = readWord.WordUtil("m2c.scm.api-1.3.0.docx")
    data=a.get_tablesdata()
    b=writeExcel(data)
    b.write_tabledata()
