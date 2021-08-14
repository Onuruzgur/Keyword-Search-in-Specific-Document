
from os import path
from PyQt5 import (QtWidgets, uic, QtCore, QtGui)
from PyQt5.QtWidgets import *
import sys
import numpy as np
import pandas as pd
import xlsxwriter
import  datetime
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMessageBox
class main(QMainWindow):
    def __init__(self):
        super(main,self).__init__()
        uic.loadUi("main.ui",self)
        self.setWindowTitle("Redstar - EASA biweekly control")
        width = 715
        height = 350
        self.setMaximumSize(width, height)
        self.browse.clicked.connect(self.browsefiles)
        self.okButton.clicked.connect(self.matching)
        self.warning.hide()
        
        

    def browsefiles(self):
        fname = QFileDialog.getOpenFileName(self,'Open file',r'C:\Users')
        self.filepath.setText(fname[0])
        
    
    def matching(self):
        
        path = self.filepath.text()
        if path.endswith('.html'):
            data = pd.read_html(str(path))
            table = data[0].to_numpy()
            info = data[1].to_numpy()
            m,n = table.shape
            null = np.zeros((m,1))
            table_append = np.append(table,null,axis=1)

            keywords = self.keywords.text()
            
            keywords_split = keywords.split(",")
            Tc_holder = keywords_split[0]
            Type = keywords_split[1]
            
        
            for i in range(m):
                k=2
                string =Tc_holder
                if string in table[i][k] :
                    x=3
                    if Type in table[i][x] :
                        table_append[i][n]= "!!!"
                        
                    else:
                        table_append[i][n]= "N/A by Type"
                    
                else:
                    table_append[i][n]= "N/A by TC Holder"

            for i in range(m):
                k=2
                string_1 = "APPLIANCES"
                if table[i][k] == string_1:
                    table_append[i][n]= "!!!"
                            
            workbook = xlsxwriter.Workbook('easa_biweekly.xlsx')
            worksheet = workbook.add_worksheet()

            row =0 
            col = 0
            cell_format = workbook.add_format()
            cell_format.set_bold()
            cell_format.set_font_color('red')

            worksheet.write('A1', 'Publ. number')
            worksheet.write('B1', 'Issue date')
            worksheet.write('C1', 'TC Holder')
            worksheet.write('D1', 'Type')
            worksheet.write('E1', 'Subject')
            worksheet.write('F1', 'Comment')

            for i in range(1,m+1):
                for k in range(n+1):
                    if k == n:
                        worksheet.write(i,k,table_append[i-1][k],cell_format)
                    else:
                        worksheet.write(i,k,table_append[i-1][k])
                
            worksheet.write(m+5,1,"Kenan ASLAN",cell_format)
                
            now = datetime.now() # current date and time
            date_time = now.strftime("%m/%d/%Y")
            worksheet.write(m+5,5,date_time,cell_format)


            workbook.close()            

            
        else:
            app = QApplication(sys.argv)
            msg = QMessageBox()
            msg.setWindowTitle("Warning")
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Please choose a html file.")
            msg.exec_()
            
   

app = QApplication(sys.argv)
Main = main()
Main.show()
app.exec()


