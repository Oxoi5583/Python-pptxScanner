"""
Create a GUI object
"""
from PyQt5.QtWidgets import QApplication,QWidget,QLabel ,QPushButton,QGridLayout,QLineEdit,QSpacerItem,QSizePolicy,QCheckBox,QMessageBox,QFileDialog,QVBoxLayout,QScrollArea
from PyQt5.QtGui import QFont,QIntValidator
from PyQt5.QtCore import Qt
import sys
from math import floor
from screeninfo import get_monitors
import threading

from Lib.PPtx_Scanner import *

def GetResolution():
    for monitor in get_monitors():
        return [monitor.width, monitor.height]


class AskforSourceFile (QWidget):
    def __init__(self,Resolution_list=GetResolution()):
        super().__init__()
        self.Resolution_list = Resolution_list
        self.setWindowTitle("Choose a .pptx file")
        self.setWindowIconText("Choose a .pptx file")
        Size_x = 250
        Size_y = 120
        self.setGeometry(floor(self.Resolution_list[0]/2-Size_x/2),floor(self.Resolution_list[1]/2-Size_y/2),Size_x,Size_y)
        self.setFixedSize(Size_x,Size_y)
        self.initUI()

    def initUI(self):
        

        gridlayout = QGridLayout()
        self.setLayout(gridlayout)


        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(10)


        self.mylabel = QLabel("Source of .pptx :",self)
        #self.mylabel.move(40,50)
        self.mylabel.setFont(font)
        gridlayout.addWidget(self.mylabel,0,0)

        self.PPTX_Source = QLineEdit()
        gridlayout.addWidget(self.PPTX_Source,1,0)
        self.SearchinDir  = QPushButton(">")
        self.SearchinDir.clicked.connect(self.SearchDir)
        gridlayout.addWidget(self.SearchinDir,1,1)


        self.Importpptx  = QPushButton("Import")
        self.Importpptx.clicked.connect(self.ImportPPTX)
        gridlayout.addWidget(self.Importpptx,2,0)
    
    def SearchDir(self):   
        fname = QFileDialog.getOpenFileName(self,"Open File", "","ppt (*.ppt *.pptx);;All Files (*)")
        self.PPTX_Source.setText(fname[0])
        
    def ImportPPTX(self):  
        PPTX_Source[0] = self.PPTX_Source.text()
        if PPTX_Source[0] != "":
            self.close()
        


##Main Window

class MainWindow (QWidget):
    def __init__(self,Resolution_list=GetResolution()):
        super().__init__()
        self.Resolution_list = Resolution_list
        self.TextFrames = []
        self.Target_slide = [0]
        self.Target_shape = [0]
        self.Target_para = [0]
        self.Target_run = [0]
        #self.onCheckboxChange()
        self.PPTX_PAGE = PPtx_Page(PPTX_Source[0])
        self.startpage = 0
        self.endpage = 14
        self.setWindowTitle("Choose a .pptx file")
        self.setWindowIconText("Choose a .pptx file")
        Size_x = 1250
        Size_y = 900
        self.setGeometry(floor(self.Resolution_list[0]/2-Size_x/2),floor(self.Resolution_list[1]/2-Size_y/2),Size_x,Size_y)
        self.setFixedSize(Size_x,Size_y)
        self.initUI()

    def initUI(self):
        
        verticallayout = QVBoxLayout()
        
        gridlayout = QGridLayout()
        gridlayout_page = QGridLayout()
        verticallayout.addLayout(gridlayout,0)
        verticallayout.addLayout(gridlayout_page,1)
        self.setLayout(verticallayout)


        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(10)

        self.scrollarea = QScrollArea()
        self.mylabel = QLabel("",self)
        self.mylabel.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scrollarea.setWidget(self.mylabel)
        self.scrollarea.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scrollarea.setWidgetResizable(True)
        #self.mylabel.move(40,50)
        self.mylabel.setFont(font)
        gridlayout.addWidget(self.scrollarea,0,0)
        """
        self.mylineedit = QLineEdit(self)
        #self.mylineedit.setMaximumSize(100,100)
        gridlayout.addWidget(self.mylineedit,0,1)
        """


        spacer = QSpacerItem(0, 0)
        gridlayout.addItem(spacer,5,0)



        #print("########")
        #print(self.startpage)
        #print(self.endpage)


        self.lastpage = QPushButton("<",self)
        self.lastpage.clicked.connect(self.onButtonClick_lastpage)
        self.lastpage.setMaximumWidth(50)
        self.lastpage.setMinimumWidth(50)
        gridlayout_page.addWidget(self.lastpage,6,0)

        #spacer = QSpacerItem(999, 0)
        #gridlayout.addItem(spacer,6,1)

        ##Page 0
        button = QPushButton(str(0),self)
        self.page0= button
        self.page0.clicked.connect(lambda:self.CheckPage(int(self.page0.text())))
        gridlayout_page.addWidget(self.page0,6,2)
        
        ##Page 1
        button = QPushButton(str(1),self)
        self.page1= button
        self.page1.clicked.connect(lambda:self.CheckPage(int(self.page1.text())))
        gridlayout_page.addWidget(self.page1,6,3)

        ##Page 2
        button = QPushButton(str(2),self)
        self.page2= button
        self.page2.clicked.connect(lambda:self.CheckPage(int(self.page2.text())))
        gridlayout_page.addWidget(self.page2,6,4)

        ##Page 3
        button = QPushButton(str(3),self)
        self.page3= button
        self.page3.clicked.connect(lambda:self.CheckPage(int(self.page3.text())))
        gridlayout_page.addWidget(self.page3,6,5)

        ##Page 4
        button = QPushButton(str(4),self)
        self.page4= button
        self.page4.clicked.connect(lambda:self.CheckPage(int(self.page4.text())))
        gridlayout_page.addWidget(self.page4,6,6)

        ##Page 5
        button = QPushButton(str(5),self)
        self.page5= button
        self.page5.clicked.connect(lambda:self.CheckPage(int(self.page5.text())))
        gridlayout_page.addWidget(self.page5,6,7)

        ##Page 6
        button = QPushButton(str(6),self)
        self.page6= button
        self.page6.clicked.connect(lambda:self.CheckPage(int(self.page6.text())))
        gridlayout_page.addWidget(self.page6,6,8)

        ##Page 7
        button = QPushButton(str(7),self)
        self.page7= button
        self.page7.clicked.connect(lambda:self.CheckPage(int(self.page7.text())))
        gridlayout_page.addWidget(self.page7,6,9)

        ##Page 8
        button = QPushButton(str(8),self)
        self.page8= button
        self.page8.clicked.connect(lambda:self.CheckPage(int(self.page8.text())))
        gridlayout_page.addWidget(self.page8,6,10)

        ##Page 9
        button = QPushButton(str(9),self)
        self.page9= button
        self.page9.clicked.connect(lambda:self.CheckPage(int(self.page9.text())))
        gridlayout_page.addWidget(self.page9,6,11)

        ##Page 10
        button = QPushButton(str(10),self)
        self.page10= button
        self.page10.clicked.connect(lambda:self.CheckPage(int(self.page10.text())))
        gridlayout_page.addWidget(self.page10,6,12)

        ##Page 11
        button = QPushButton(str(11),self)
        self.page11= button
        self.page11.clicked.connect(lambda:self.CheckPage(int(self.page11.text())))
        gridlayout_page.addWidget(self.page11,6,13)

        ##Page 12
        button = QPushButton(str(12),self)
        self.page12= button
        self.page12.clicked.connect(lambda:self.CheckPage(int(self.page12.text())))
        gridlayout_page.addWidget(self.page12,6,14)

        ##Page 13
        button = QPushButton(str(13),self)
        self.page13= button
        self.page13.clicked.connect(lambda:self.CheckPage(int(self.page13.text())))
        gridlayout_page.addWidget(self.page13,6,15)
        
        spacer = QSpacerItem(999, 0)
        gridlayout_page.addItem(spacer,6,16)

        self.nextpage = QPushButton(">",self)
        self.nextpage.clicked.connect(self.onButtonClick_nextpage)
        self.nextpage.setMaximumWidth(50)
        self.nextpage.setMinimumWidth(50)
        gridlayout_page.addWidget(self.nextpage,6,17)


        self.pages = [
        self.page0,
        self.page1,
        self.page2,
        self.page3,
        self.page4,
        self.page5,
        self.page6,
        self.page7,
        self.page8,
        self.page9,
        self.page10,
        self.page11,
        self.page12,
        self.page13]
        for page in self.pages:
            page.setMaximumWidth(100)
            print("##")
            if int(page.text()) >= self.PPTX_PAGE:
                page.hide()
            else: 
                page.show()

    def CheckPage(self,i):
        print(i)
        Text = ""
        if int(i) < int(self.PPTX_PAGE):
            slide = int(i)
            self.TextFrames = PPtx_TextFrame(Path = PPTX_Source[0],slide=slide)
            print(self.TextFrames)
            
            for shape in self.TextFrames:
                shape_list = shape.split(":")
                print(shape_list)
                self.Target_slide[0] = int(shape_list[0])
                self.Target_shape[0] = int(shape_list[1])
                self.Target_para[0] = int(shape_list[2])
                self.Target_run[0]= int(shape_list[3])
                Text = Text +  "slide[{}].shapes[{}].text_frame.paragraphs[{}].runs[{}]".format(shape_list[0],shape_list[1],shape_list[2],shape_list[3]) + "\n"
                Text = Text + PPtx_GetText(PPTX_Source[0],self.Target_slide[0],self.Target_shape[0],self.Target_para[0],self.Target_run[0]) +"\n\n"
            print(PPTX_Source[0])
            #self.mylabel.setText(str(self.TextFrames))
            self.mylabel.setText(Text)
            

    def onButtonClick_lastpage(self):
        print("Last page clicked")
        if int(self.page0.text()) > 0:
            self.page0.setText(str(int(self.page0.text())-14))
            self.page1.setText(str(int(self.page1.text())-14))
            self.page2.setText(str(int(self.page2.text())-14))
            self.page3.setText(str(int(self.page3.text())-14))
            self.page4.setText(str(int(self.page4.text())-14))
            self.page5.setText(str(int(self.page5.text())-14))
            self.page6.setText(str(int(self.page6.text())-14))
            self.page7.setText(str(int(self.page7.text())-14))
            self.page8.setText(str(int(self.page8.text())-14))
            self.page9.setText(str(int(self.page9.text())-14))
            self.page10.setText(str(int(self.page10.text())-14))
            self.page11.setText(str(int(self.page11.text())-14))
            self.page12.setText(str(int(self.page12.text())-14))
            self.page13.setText(str(int(self.page13.text())-14))
        
        
        for page in self.pages:
            print("##")
            if int(page.text()) >= self.PPTX_PAGE:
                page.hide()
            else: 
                page.show()


    def onButtonClick_nextpage(self):
        print("Next page clicked")
        


        if int(self.page13.text()) < self.PPTX_PAGE:
            self.page0.setText(str(int(self.page0.text())+14))
            self.page1.setText(str(int(self.page1.text())+14))
            self.page2.setText(str(int(self.page2.text())+14))
            self.page3.setText(str(int(self.page3.text())+14))
            self.page4.setText(str(int(self.page4.text())+14))
            self.page5.setText(str(int(self.page5.text())+14))
            self.page6.setText(str(int(self.page6.text())+14))
            self.page7.setText(str(int(self.page7.text())+14))
            self.page8.setText(str(int(self.page8.text())+14))
            self.page9.setText(str(int(self.page9.text())+14))
            self.page10.setText(str(int(self.page10.text())+14))
            self.page11.setText(str(int(self.page11.text())+14))
            self.page12.setText(str(int(self.page12.text())+14))
            self.page13.setText(str(int(self.page13.text())+14))

        for page in self.pages:
            print("##")
            if int(page.text()) >= self.PPTX_PAGE:
                page.hide()
            else: 
                page.show()
                
    """   
    def onCheckboxChange(self):
        if self.IsCombinedNeeded.isChecked():
            self.mylineedit2.setHidden(False)
            self.mylabel2.setHidden(False)
            self.Hint2.setHidden(False)
            Size_x = 250
            Size_y = 180
            self.setGeometry(floor(self.Resolution_list[0]/2-Size_x/2),floor(self.Resolution_list[1]/2-Size_y/2),Size_x,Size_y)
            self.setFixedSize(Size_x,Size_y)
        else:
            self.mylineedit2.setHidden(True)
            self.mylabel2.setHidden(True)
            self.Hint2.setHidden(True)
            Size_x = 250
            Size_y = 130
            self.setGeometry(floor(self.Resolution_list[0]/2-Size_x/2),floor(self.Resolution_list[1]/2-Size_y/2),Size_x,Size_y)
            self.setFixedSize(Size_x,Size_y)
"""

if __name__ == "__main__":
    #print(GetResolution())
    PPTX_Source = [""]

    app = QApplication(sys.argv)
    First_w = AskforSourceFile()
    
    First_w.show()
    app.exec_()
    print(PPTX_Source[0])
    if PPTX_Source[0] != "":
        #w.initUI()
        Main_w = MainWindow()
        Main_w.show()
        sys.exit(app.exec_())

