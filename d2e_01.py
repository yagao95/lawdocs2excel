# -*- coding: UTF-8 -*-
"""
1.加载一个指定路径文件夹内的所有txt内容
2.把解析出来的指定内容写入Excel表格
"""
import docx
import xlrd
import xlwt
from xlutils.copy import copy
import os
import re
import threading
import time
import shutil
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QLabel, QHBoxLayout, QVBoxLayout, qApp, QDesktopWidget, QFileDialog, QPlainTextEdit

__author__ = "yooongchun" 
__email__ = "yooongchun@foxmail.com"
# #以上为原作者信息，2018-12-31,fangtao95为适用法律文书而改编，邮箱：fangtao6793@163.com


# 加载某文件夹下的所有TXT文件，返回其绝对路径
def loadTXT(file_path):
    txt_files = []  # 保存文件地址和名称
    files = os.listdir(file_path)
    for _file in files:
        if not os.path.splitext(_file)[1] == '.docx':  # 判断是否为docx文件
            continue
        abso_path = os.path.join(file_path, _file)
        txt_files.append(abso_path)
    return txt_files


# 提取TXT文件内容
def extractor(txt_path):
    doc = docx.Document(txt_path)
    fullText = []
    for i in doc.paragraphs:  # 迭代docx文档里面的每一个段落
        fullText.append(i.text)  # 保存每一个段落的文本
    text = '\n'.join(fullText)
    text = re.sub(r":", "：", text)
    text = re.sub(r",", "，", text)
    text = re.sub(r"\u3000", "", text)

#  #############################定义变量
    yiju = ""
    anyou = ""
    anqing = ""
    wenhao = ""
    #以上变量的value将为[]
    fayuan_name = ""
    pancaishu_name = ""
    panjue = ""
    qingqiu = ""
    jiaodian = ""
#  ############################定义列表[]和字典{}
    beigao = []
    yuangao = []
    shenpanyuan = []
    peishenyaun = []

    fayuan_name1 = []
    pancaishu_name1 = []
    panjue1 = []
    qingqiu1 = []
    jiaodian1 = []
    falv1 = []
    fatiao1 = []
    tiaotexts1 = []

    shenpanzhang = []
    zhuli = []
    shuji = []
    jingyingzhe = []
    disanren = []
    
    INFO = []  # * *******************变量INFO=list索引/列表********************************************
    info = {}  # ***************变量info=dict字典*************************************************

    headers = text.split("\n")


#  ***************header查找******
    for index, one in enumerate(headers):
        header = one.split(" ")
        
        key = "审判长"  
        if key in one: 
            v = re.sub(r"\s+", " ", one)  
            v = v.split("审判长")[-1]           
            shenpanzhang.append(v)
            info[key] = shenpanzhang                            
            continue

        # 查找审判员
        key = "审判员" 
        if key in one:            
            v = re.sub(r"\s+", " ", one) 
            v = v.split("审判员")[-1]
            shenpanyuan.append(v)
            info[key] = shenpanyuan                 
            continue  

        # 查找人民陪审员
        key = "人民陪审员"  
        if key in one:             
            v = re.sub(r"\s+", " ", one)  
            v = v.split("人民陪审员")[-1]           
            peishenyaun.append(v)
            info[key] = peishenyaun                 
            continue 

        # 查找法官助理
        key = "法官助理" 
        if key in one:          
            v = re.sub(r"\s+", " ", one)  
            v = v.split("法官助理")[-1]          
            zhuli.append(v)
            info[key] = zhuli                 
            continue 

        # 查找书记员
        key = "书记员"
        if key in one:           
            v = re.sub(r"\s+", " ", one)  
            v = v.split("书记员")[-1]          
            shuji.append(v)
            info[key] = shuji                 
            continue  

 #  ################################################################
    for index, one in enumerate(headers):
            if re.sub(r"\s*", "", one) == "民事判决书":
                    fayuan_name = headers[index - 1]
                    fayuan_name = re.sub(r"\s+", "", fayuan_name)
                    pancaishu_name = "民事判决书"
                    break
            elif re.sub(r"\s*", "", one) == "民事裁定书":
                    fayuan_name = headers[index - 1]
                    fayuan_name = re.sub(r"\s+", "", fayuan_name)
                    pancaishu_name = "民事裁定书"
                    break

    for index, one in enumerate(headers):
            if re.findall(r"依照(.*)规定", one):
                    yiju = re.findall(r"依照(.*)规定", one)


    text1 = re.sub(r"\s+", "", text)            
    text2 = re.split(r"\s+|[，,:：.。;；]+", text1)
    text3 = re.split(r"\s+|[。]+", text1) 
    
    for index, one in enumerate(text2):  #  对于text2中的每个元素分别取出，且命名为index、one
            if re.findall(r"法院民事..书(.*)原告", one):                           
                    wenhao = re.findall(r"法院民事..书(.*)原告", one)
            if re.findall(r"原告.+被告.+纠纷一案", one):               
                    anqing = re.findall(r"原告.+被告.+纠纷一案", one)
                    anyou = re.findall(r"垄断纠纷|垄断协议纠纷|横向垄断协议纠纷|纵向垄断协议纠|滥用市场支配地位纠纷|垄断定价纠纷|掠夺定价纠纷|拒绝交易纠纷|限定交易纠纷|捆绑交易纠纷|差别待遇纠纷|经营者集中纠纷", one)
                    
    for index, one in enumerate(text3):
            if re.findall(r"原告.+被告.+纠纷一案.+立案", one):               
                    anqing = re.findall(r"原告.+被告.+纠纷一案.+立案", one)
                    anyou = re.findall(r"垄断纠纷|垄断协议纠纷|横向垄断协议纠纷|纵向垄断协议纠纷|滥用市场支配地位纠纷|垄断定价纠纷|掠夺定价纠纷|拒绝交易纠纷|限定交易纠纷|捆绑交易纠纷|差别待遇纠纷|经营者集中纠纷", one)
                    
           #  查找诉讼请求
            if "诉讼请求：" in one: 
                    qingqiu = re.sub(r"\s+", " ", one)  
                    qingqiu = qingqiu.split("：")[-1]
                    continue
            elif "判令：" in one: 
                    qingqiu = re.sub(r"\s+", " ", one)  
                    qingqiu = qingqiu.split("：")[-1]
                    continue
            #  查找焦点问题
            if "本案的焦点问题" in one: 
                    jiaodian = re.sub(r"\s+", " ", one)  
                    jiaodian = jiaodian.split("：")[-1]
                    continue
           #  查找裁判结果
            if "裁定如下：" in one: 
                    panjue = re.sub(r"\s+", " ", one)  
                    panjue = panjue.split("：")[-1]
                    continue
            elif "判决如下：" in one: 
                    panjue = re.sub(r"\s+", " ", one)  
                    panjue = panjue.split("：")[-1]
                    continue

#  ############################################查找法律、法条
    if len(yiju) != 0:
            
        yijutext1 = yiju[0]
        yijutext2 = re.split(r"，|,|、", yijutext1)   # 以符号分割




        for index, one in enumerate(yijutext2): 
                if re.findall(r"《.*》", one):                          
                        falv1 = re.findall(r"《.*》", one)

        if len(yijutext2) > 1:
                yijutext3 = re.sub(r"之", "", yijutext1)
                yijutext3 = re.split(r"，|,|《", yijutext3)
                for index, one in enumerate(yijutext3):
                        if re.findall(r".*》第.+条", one):              
                                fatiao = re.findall(r".*》第.+条", one)
                                fatiao1.append(fatiao)
                                fatiao1.append(fatiao)                                
                                tiaohead = one.split("》")[0]
                                tiaotexts = one.split("》")[1]
                                tiaotext1 = tiaotexts.split("、")
                                for i,p in enumerate(tiaotext1):
                                        tiaotext = '《'+ tiaohead + '》' + p
                                        tiaotexts1.append(tiaotext)
        else:
                tiaotexts1.append(yijutext1)
        
#  ##########

            
    info["案情"] = anqing
    info["案由"] = anyou
    info["裁判依据"] = yiju
    info["案件字号"] = wenhao  
    #  以上value是[]


    fayuan_name1.append(fayuan_name)
    pancaishu_name1.append(pancaishu_name)
    panjue1.append(panjue)
    qingqiu1.append(qingqiu)
    jiaodian1.append(jiaodian)
    
    info["裁判结果"] = panjue1                
    info["审理法院"] = fayuan_name1   
    info["文书类型"] = pancaishu_name1  
    info["本案的焦点问题"] = jiaodian1  
    info["诉讼请求"] = qingqiu1
    
    info["法律依据"] = falv1  
    info["法条依据"] = tiaotexts1

    
#  ##########
    disclosure = re.split(r"\s+|[，。]+", text)    
    counter = -1    
    for one in disclosure[1:]:    
        counter += 1
        text = re.split(r"\n", one)  
        for index, item in enumerate(text):
            cc = -1
            
    #  查找4原告
            cc += 1
            key = "原告："  # 关键词key="原告"
            if key in item:  # 如果关键词在item
                # flag[cc] = False
                v = re.sub(r"\s+", " ", item)  # 变量v=在item中删去\s换行符等
                v = v.split("：")[-1]           # 变量v=删去“：”后的后一个v
                yuangao.append(v)
                info[key] = yuangao                  # 将变量v的值以key名加入字典info中
                continue  # 如果关键词在item中则继续循环,直到没有此关键词。

    # 查找5被告
            cc += 1
            key = "被告：" 
            if key in item: 
                v = re.sub(r"\s+", " ", item)  
                v = v.split("：")[-1]           
                beigao.append(v)
                info[key] = beigao                 
                continue   # 如果关键词在item中则继续循环,直到没有此关键词。

    # 查找14经营者
            cc += 1
            key = "经营者：" 
            if key in item: 
                v = re.sub(r"\s+", " ", item)  
                v = v.split("：")[-1]           
                jingyingzhe.append(v)
                info[key] = jingyingzhe                  
                continue
    # 查找15第三人
            cc += 1
            key = "第三人：" 
            if key in item: 
                v = re.sub(r"\s+", " ", item)  
                v = v.split("：")[-1]           
                disanren.append(v)
                info[key] = disanren                  
                continue

    INFO.append(info)
    # 将字典info添加到list列表INFO中
    return INFO


# 为每个PDF文件添加sheet
def add_sheet(Excel_path, names):
    book = xlrd.open_workbook(Excel_path)  # 打开一个wordbook
    sheet = book.sheet_by_name("output")  # 打开名为output的sheet
    key_words = sheet.row_values(1, 0, sheet.ncols)  # key_words = sheet行值.row_values
    book = xlwt.Workbook()
    for name in names:
        sheet = book.add_sheet(name, cell_overwrite_ok=True)
        for i in range(len(key_words)):
            sheet.write(0, i, key_words[i])
    new_path = os.path.splitext(Excel_path)[0] + "_Result.xls"
    book.save(new_path)
    return new_path


# 保存到Excel中
def save2Excel(INFO, Excel_path, sheet_name):
    book = xlrd.open_workbook(Excel_path)  # 打开一个wordbook
    sheet = book.sheet_by_name(sheet_name)
    key_words = sheet.row_values(0, 0, sheet.ncols)
    copy_book = copy(book)
    sheet_copy = copy_book.get_sheet(sheet_name)

#    info["源文件名"] = sheet_name 不可在此，否则程序崩溃

    for index, info in enumerate(INFO):
        sheet_name1 = []
        sheet_name1.append(sheet_name)
        info["源文件名"] = sheet_name1
        for key, value in info.items():
            for ind, key_out in enumerate(key_words):
                if key_out == key:
                    if len(value)== 1:
                        col = ind
                        row = index + 1
                        sheet_copy.write(row, col, value)
                        break
                    else:
                        col = ind
                        row = index + 1
                        i = row
                        i = i - 1
                        for j,q in enumerate(value):
                            i = i + 1
                            sheet_copy.write(i,col,q)
                            
                    break
                
    copy_book.save(Excel_path)


# GUI界面代码
class MYGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.exit_flag = False
        self.try_time = 1532052770.2892158 + 24 * 60 * 60  # 试用时间

        self.initUI()

    def initUI(self):
        self.pdf_label = QLabel("docx文件夹路径: ")
        self.pdf_btn = QPushButton("选择")
        self.pdf_btn.clicked.connect(self.open_pdf)
        self.pdf_path = QLineEdit("docx文件夹路径...")
        self.pdf_path.setEnabled(False)
        self.excel_label = QLabel("Excel Demo 路径: ")
        self.excel_btn = QPushButton("选择")
        self.excel_btn.clicked.connect(self.open_excel)
        self.excel_path = QLineEdit("Excel Demo路径...")
        self.excel_path.setEnabled(False)
        self.output_label = QLabel("保存路径: ")
        self.output_path = QLineEdit("保存文件路径...")
        self.output_path.setEnabled(False)
        self.output_btn = QPushButton("选择")
        self.output_btn.clicked.connect(self.open_output)
        self.info = QPlainTextEdit()

        h1 = QHBoxLayout()
        h1.addWidget(self.pdf_label)
        h1.addWidget(self.pdf_path)
        h1.addWidget(self.pdf_btn)

        h2 = QHBoxLayout()
        h2.addWidget(self.excel_label)
        h2.addWidget(self.excel_path)
        h2.addWidget(self.excel_btn)

        h3 = QHBoxLayout()
        h3.addWidget(self.output_label)
        h3.addWidget(self.output_path)
        h3.addWidget(self.output_btn)

        self.run_btn = QPushButton("运行")
        self.run_btn.clicked.connect(self.run)

        self.auth_label = QLabel("作者邮箱")
        self.auth_ed = QLineEdit("fangtao6793@163.com")

        exit_btn = QPushButton("退出")
        exit_btn.clicked.connect(self.Exit)
        h4 = QHBoxLayout()
        h4.addWidget(self.auth_label)
        h4.addWidget(self.auth_ed)
        h4.addStretch(1)
        h4.addWidget(self.run_btn)
        h4.addWidget(exit_btn)

        v = QVBoxLayout()
        v.addLayout(h1)
        v.addLayout(h2)
        v.addLayout(h3)
        v.addWidget(self.info)
        v.addLayout(h4)
        self.setLayout(v)
        width = int(QDesktopWidget().screenGeometry().width() / 3)
        height = int(QDesktopWidget().screenGeometry().height() / 3)
        self.setGeometry(100, 100, width, height)
        self.setWindowTitle('Docx to Excel')
        self.show()

    def Exit(self):
        self.exit_flag = True
        qApp.quit()

    def open_pdf(self):
        fname = QFileDialog.getExistingDirectory(self, "Open folder", "/home")
        if fname:
            self.pdf_path.setText(fname)

    def open_excel(self):
        fname = QFileDialog.getOpenFileName(self, "Open Excel", "/home")
        if fname[0]:
            self.excel_path.setText(fname[0])

    def open_output(self):
        fname = QFileDialog.getExistingDirectory(self, "Open output folder",
                                                 "/home")
        if fname:
            self.output_path.setText(fname)

    def run(self):
        self.info.setPlainText("")
        threading.Thread(target=self.scb, args=()).start()
        if self.auth_ed.text() == "fangtao6793@163.com":
            self.info.insertPlainText("密码正确，开始运行程序!\n")
            threading.Thread(target=self.main_fcn, args=()).start()
        elif self.auth_ed.text() == "test_mode":
            if time.time() < self.try_time:
                self.info.insertPlainText("试用模式，截止时间：2018-07-22 10:00\n")
                threading.Thread(target=self.main_fcn, args=()).start()
            else:
                self.info.insertPlainText(
                    "试用已结束，继续使用请联系fangtao6793@163.com获取密码\n")

        else:
            self.info.insertPlainText(
                "密码错误，请联系fangtao6793@163.com获取正确密码!\n")

    def scb(self):
        flag = True
        cnt = self.info.document().lineCount()
        while not self.exit_flag:
            if flag:
                self.info.verticalScrollBar().setSliderPosition(
                    self.info.verticalScrollBar().maximum())
            time.sleep(0.01)
            if cnt < self.info.document().lineCount():
                flag = True
                cnt = self.info.document().lineCount()
            else:
                flag = False
            time.sleep(0.01)

    def main_fcn(self):
        # 加载docx文件夹
        if os.path.isdir(self.pdf_path.text()):
            try:
                txt_path = self.pdf_path.text()
            except Exception:
                self.info.insertPlainText("加载docx文件夹出错，请重试！\n")
                return
        else:
            self.info.insertPlainText("docx路径错误，请重试！\n")
            return
        # 加载Excel路径
        if os.path.isfile(self.excel_path.text()):
            demo_path = self.excel_path.text()
        else:
            self.info.insertPlainText("Excel路径错误，请重试！\n")
            return
        # 加载保存路径
        if os.path.isdir(self.output_path.text()):
            name = os.path.basename(demo_path)
            out_path = os.path.join(self.output_path.text(),
                                    name.replace(".xlsx", ".xls"))
        else:
            self.info.insertPlainText("输出路径错误，请重试！\n")
            return
        try:
            shutil.copyfile(demo_path, out_path)
        except Exception as e:
            traceback.print_exc()
            self.info.insertPlainText("拷贝临时文件出错，请确保程序有足够运行权限再重试！可能有文件处于正在被打开状态\n")
            return
        try:
            self.info.insertPlainText("加载docx文件...\n")
            txt_paths = loadTXT(txt_path)
        except Exception as e:
            traceback.print_exc()
            self.info.insertPlainText("加载docx文件出错！\n")
            return
        names = [os.path.basename(name).split(".")[0] for name in txt_paths]
        try:
            new_path = add_sheet(out_path, names)
        except Exception as e:
            traceback.print_exc()
            self.info.insertPlainText("生成sheet失败！\n")
            return
        counter = 0
        for name, path in zip(names, txt_paths):
            counter += 1
            self.info.insertPlainText("正在处理文件: %s %d/%d" %
                                      (name, counter, len(names)) + "\n")
            try:
                INFO = extractor(path)
                save2Excel(INFO, new_path, name)
            except Exception as e:
                traceback.print_exc()
                self.info.insertPlainText("文件：%s 出错，跳过...\n" % name)
                continue
        self.info.insertPlainText("运行完成！\n")


# 程序入口
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MYGUI()
    sys.exit(app.exec_())
