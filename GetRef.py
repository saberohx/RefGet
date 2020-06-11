from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup as soup
import xlsxwriter
import time
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import resources_rc
import pandas as pd
import os
import subprocess
import threading
from requests import get
MainUi,_=loadUiType('MyDesign.ui')

class Main(QWidget,MainUi):
    def __init__(self,parent=None):
        super(Main,self).__init__(parent)
        QWidget.__init__(self)
        self.setupUi(self)
        self.changeTab_home(0)
    #------------------------- Interface ------------------------------#
        self.but_search.clicked.connect(self.exec)
        self.but_home.clicked.connect(lambda:self.changeTab_home(0))
        self.but_about.clicked.connect(lambda:self.changeTab_home(2))
        self.but_settings.clicked.connect(lambda:self.changeTab_home(1))
        self.but_import.clicked.connect(self.import_file)
        self.but_browse.clicked.connect(self.output_file)
        #self.but_about.clicked.connect(lambda:self.open_file("C:/Users/hp/Desktop/llllllllll.xlsx"))
    #--------------------- Changing the tab --------------------------------------#
    def changeTab_home(self,index):
        self.tab_2.setCurrentIndex(index)


    #-------------------------- Import xlsx -------------------------------
    def import_file(self):
        try:
            global path_import
            path_import = QFileDialog.getOpenFileName(
                parent=self,
                caption='Select InPut File',filter="*.xlsx *.csv")
            #print(path_import)
            self.txt_file.setText(path_import[0])
            global df,file_size
            df = pd.read_excel(path_import[0])
            file_size=int(df.count())

            self.show_ref(str(file_size) + " References will be processed !")
        except Exception as exc:
            #print(exc)
            self.txt_warnings.setText(exc)

    #------------------------ Output Directory -------------------------------#
    def output_file(self):
        try:
            path = QFileDialog.getExistingDirectory(
                parent=self,
                caption='Select OutPut directory'
            )
            #print(path)

            self.txt_folder.setText(path)

        except Exception as exc:
            #print(exc)
            self.txt_warnings.setText(exc)
    #------------------ MsgBox --------------------------------------------#

    def check_supp(self):
        global supObj
        supObj = []
        if self.check_valeo.isChecked():
            supObj.append("VALEO")
        if self.check_sachs.isChecked():
            supObj.append("SACHS")
        if self.check_aisin.isChecked():
            supObj.append("AISIN")
        if self.check_luk.isChecked():
            supObj.append("LUK")

    #------------------ The Main Operation ------------#
    def create_xslx(self):
        path_file=self.txt_folder.text()
        file_name=self.txt_name.text()
        global directory
        directory=path_file+"/"+file_name+".xlsx"


        global ws_luk,ws_warn,ws_sashs,ws_valeo,ws_aisin,wb

        wb = xlsxwriter.Workbook(directory)
        ws_aisin = wb.add_worksheet("AISIN")
        ws_sashs = wb.add_worksheet("SASHS")
        ws_luk = wb.add_worksheet("LUK")
        ws_valeo = wb.add_worksheet("VALEO")
        ws_warn = wb.add_worksheet("WARNINGS")
        #print("xlsx file Done !")
        self.show_ref("Excel file Done !")
    def progress_check(self,processed):
        percent = 100 * processed / file_size
        self.progressBar.setValue(int(percent))
    def get(self):

        preurl = "http://web2.carparts-cat.com/default.aspx?11=426"
        chromedriver = r"C:\Users\hp\Documents\Software\chromedriver_win32\chromedriver.exe"
        # firefoxdriver=r"C:\Users\hp\Documents\Software\firfoxdriver\geckodriver.exe"
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--incognito')
        options.add_argument('--headless')
        browser = webdriver.Chrome(chromedriver, options=options)
        #print('Opening Browser ...')
        self.show_ref('Opening Browser ...')
        # cette partie pour vérifier la connection
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
            req = get(preurl,headers=headers)
            req.raise_for_status()
        except Exception as exc:
            self.show_ref('Please check your connection !')
            print('Please check your connection ! %s'%exc)
        #self.label.setText(preurl)
        browser.get(preurl)
        try:
            ##username
            WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID,"username")))
            self.show_ref('Login ...')
            #print("Login ...")
            user=browser.find_element_by_id("username")
            #//*[@id="username"]
            user.send_keys("99991userconslt")
            browser.find_element_by_name("password").send_keys("centraweb19")
            browser.find_element_by_name("login").click()
            row_valeo = 0
            row_sachs = 0
            row_aisin = 0
            row_luk = 0
            row_ = 0
            #------------- LOOOOOOOOOOOOP --------------------------
            global processed
            processed=0
            for index, row in df.iterrows():
                ref=str(row['Références'])
                col=1
                #print("La référence : ",ref)
                self.show_ref("Processing : "+ref)
                #print(listOfRef.index(ref)+1)
                WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.ID, "tp_articlesearch_txt_articleSearch")))
                browser.find_element_by_id("tp_articlesearch_txt_articleSearch").clear()
                browser.find_element_by_id("tp_articlesearch_txt_articleSearch").send_keys(ref)
                browser.find_element_by_id("tp_articlesearch_articleSearch_imgBtn").click()

                #collapse_header_image collapsed
                #//*[@id="mainpane"]/div[2]/div/div/div[3]/div[1]/img

                WebDriverWait(browser, 100).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mainpane"]/div[2]/div/div/div[3]/div[1]/img')))
                search=browser.find_element_by_xpath('//*[@id="mainpane"]/div[2]/div/div/div[3]/div[2]/input').is_displayed()
                if not search == True:
                    li=browser.find_element_by_xpath('//*[@id="mainpane"]/div[2]/div/div/div[3]/div[1]/img')
                    li.click()

                src_page=browser.page_source
                soupCat=soup(src_page,"lxml")
                cat=soupCat.find("tr", {"class": "main_artikel_panel_tr_genart colorClass5_sub1"})
                #print(cat.span.text)
                try:
                    browser.find_element_by_name(cat.span.text).find_element_by_tag_name("input").click()
                    self.show_ref(cat.span.text)


                    time.sleep(10)
                    page = browser.page_source
                    soupObject = soup(page, "lxml")
                    supList = soupObject.find_all('span', {"title": "seulement"})
                    #nbRef = soupObject.find("label", {"for": "searchArea_2"}).text
                    #print(nbRef.split(" ")[3])
                    supName=[]
                    for k in supList:
                        n = k.text
                        supName.append(n)
                    #print(supName)
                    self.check_supp()
                    #supObj=["VALEO","SACHS","AISIN","LUK"]
                    check=[]

                    for name in supObj:
                        if name in supName:
                            sup=browser.find_element_by_name(name).find_element_by_tag_name("input")
                            sup.click()
                            time.sleep(2)
                            #print(name,"has clicked")

                            check.append(name)
                        else:
                            #print(name,"Not exist")
                            pass
                except Exception as exc:
                    print(exc,ref)
                    ws_warn.write(row_, 0, ref)
                    ws_warn.write(row_, col, "ERROR")
                    row_ += 1
                    col += 1
                    continue

                #-------------- OUTPUTS ------------------------------------

                if not len(check)==0:
                    time.sleep(10)
                    page_final=browser.page_source
                    soupObjectFinal=soup(page_final,"lxml")
                    category=soupObjectFinal.find_all("tr",{"class":"main_artikel_panel_tr_genart colorClass5_sub1"})
                    sof=soupObjectFinal.find_all("div",{"class":"pnl_link_eartnr"})
                    #supn=soupObjectFinal.find_all("tr",{"class":"main_artikel_panel_tr_einspeiser colorClass5_sub2"})
                    #print(len(sof))
                    #print("category",len(category))


                    #print(sof)

                    for element in sof:
                        col = 1
                        part=element.nobr.text
                        #print(re.search('[a-zA-Z]', part))
                        if part.isupper()==True or part.islower()==True:
                            #elif len(part) == 7 or len(part) == 8:
                            ws_aisin.write(row_aisin, 0, ref)
                            ws_aisin.write(row_aisin, col, part)
                            col += 1
                            row_aisin += 1
                        elif len(part)==6:
                            ws_valeo.write(row_valeo,0,ref)
                            ws_valeo.write(row_valeo, col, part)
                            col+=1
                            row_valeo+=1
                        elif len(part)==12:
                            ws_sashs.write(row_sachs, 0, ref)
                            ws_sashs.write(row_sachs, col, part)
                            col += 1
                            row_sachs += 1
                        elif len(part)==11:
                            ws_luk.write(row_luk, 0, ref)
                            ws_luk.write(row_luk, col, part)
                            col += 1
                            row_luk += 1

                        else:
                            ws_warn.write(row_, 0, ref)
                            ws_warn.write(row_, col, part)
                            row_ += 1
                            col += 1


                else:
                    ws_warn.write(row_, 0, ref)
                    ws_warn.write(row_, col + 1, "EMPTY")
                    row_+=1

                #break
                processed += 1
                self.progress_check(processed)
            browser.close()
            wb.close()
            #print("SAVED")

            #self.finished()
            #self.popup()
            self.show_ref("Process finished")
            #print(directory)
            self.thread_excel(directory)

        except Exception as exc:
            self.txt_warnings.setText(exc)
            browser.close()
        #wb.close()

    def exec(self):
        try :
            f1=threading.Thread(target=self.create_xslx,args=(),daemon=True)
            f2 = threading.Thread(target=self.get,args=(),daemon=True)
            #fp = threading.Thread(target=self.progress_check(processed), args=(), daemon=True)
            f1.start()
            f2.start()
            #fp.start()
        except Exception as exc :
            self.txt_warnings.setText(exc)
    def show_ref(self,text):
        try:
            f3 = threading.Thread(target=self.label.setText(text), args=(),daemon=True)
            f3.start()
        except Exception as exc:
            self.txt_warnings.setText(exc)
    def open_file(self,FileName):
        msg=QMessageBox()
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("Ask Permission")
        #msg.setWindowIcon("excel.png")
        msg.setText("Do you want to open the file ?")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.Yes)
        buttonYes = msg.button(QMessageBox.Yes)
        buttonYes.setText("Yes")
        buttonNo = msg.button(QMessageBox.No)
        buttonNo.setText("NO")
        msg.exec_()
        if msg.clickedButton() == buttonYes:
            pass
    def process_excel(self,FileName):
        subprocess.run(FileName,shell=True)
    def thread_excel(self,directory):
        try :
            #f5=threading.Thread(target=self.label.setText("Process Finished !"))
            f4=threading.Thread(target=self.process_excel(directory),args=(),daemon=True)
            f4.start()
            #f4.start()
        except Exception as exc:
            self.txt_warnings.setText(exc)
def main():
    app=QApplication(sys.argv)
    window=Main()
    window.show()
    app.exec_()
if __name__ == '__main__':
    main()
