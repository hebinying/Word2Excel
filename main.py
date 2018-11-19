#coding=utf-8
from common import readWord,writeExcel
import os
#获取需要转的文档
def get_file_list(filepath):
    dirs=os.listdir(filepath)
    list=[]
    for dir in dirs:
        filepath1 = os.path.join(filepath, dir)
        # print filepath1.decode("gbk")
        if os.path.isdir(dir):
            list1=get_file_list(filepath1)
            list.extend(list1)
        else:
            #获取docx文件
            if "docx" in os.path.splitext(dir)[1] and "~$" not in dir:
                list.append(filepath1)
    return list
#转文件
def change_file(filepath):

    # filepath=filepath.encode("gbk")
    print "转文件%s开始" % filepath.decode("gbk")
    rw = readWord.WordUtil(filepath)
    data = rw.get_tablesdata()
    # rw.close_word()
    we = writeExcel.writeExcel(data)
    we.write_tabledata()
    print "转文件%s成功" % filepath.decode("gbk")



if __name__=="__main__":
    #文件地址输入(目录获取所有docx文件并返回)
    #文件名输入
    while True:
        path=raw_input("请输入需要转换的word文件：")
        filepath=path.decode("utf-8").encode("gbk")
        if os.path.isdir(filepath):
            dirs=get_file_list(filepath)
            print dirs
            for dir in dirs:
                print dir.decode("gbk")
                change_file(dir)
        elif os.path.isfile(filepath):
            if "docx" in os.path.splitext(filepath)[1]:
                change_file(filepath)
                print "转文件"
            else:
                print "输入文件不是docx格式文件，请重试"
        else:
            print "输入的路径是错误的，请检查"





