from PyQt5.QtWidgets import QFrame,QApplication,QWidget,QLabel ,QPushButton,QGridLayout,QLineEdit,QSpacerItem,QSizePolicy,QCheckBox,QMessageBox,QFileDialog,QVBoxLayout,QScrollArea
from PyQt5.QtGui import QFont,QIntValidator,QTextBlock
from PyQt5.QtCore import Qt,QSize
from Lib.PPtx_Scanner import PPtx_Scanner,PPtx_GetText,PPtx_Page,PPtx_TextFrame


def CheckPage(PPTX_Source,self,i):
        print(i)
        for widg in self.ContentDisplay:
            widg.deleteLater()
        for spa in self.ContentDisplay_spacer:
            self.labelsLayout.removeItem(spa)
        self.ContentDisplay_spacer = []
        self.ContentDisplay = []
        if int(i) < int(self.PPTX_PAGE):
            slide = int(i)
            self.TextFrames = PPtx_TextFrame(Path=PPTX_Source,slide=slide)
            print(self.TextFrames)
            y = 0
            for shape in self.TextFrames:
                shape_list = shape.split(":")
                print(shape_list)
                Target_slide = int(shape_list[0])
                Target_shape = int(shape_list[1])
                Target_para = int(shape_list[2])
                Target_run= int(shape_list[3])
                PPtx_Text_Path = "slide[{}].shapes[{}].text_frame.paragraphs[{}].runs[{}].text".format(shape_list[0],shape_list[1],shape_list[2],shape_list[3])
                PPtx_Text = PPtx_GetText(PPTX_Source,Target_slide,Target_shape,Target_para,Target_run)

                vbox = QVBoxLayout()
                widget = QWidget()
                widget.setLayout(vbox)
                self.PPtx_Text_Path_label = QLabel(PPtx_Text_Path,self)
                self.PPtx_Text_Path_label.setAlignment(Qt.AlignmentFlag.AlignVCenter)
                self.PPtx_Text_Path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                self.PPtx_Text_Path_label.setMinimumSize(QSize(1000,50))
                self.PPtx_Text_Path_label.setMaximumSize(QSize(1000,200))
                self.PPtx_Text_Path_label.setSizePolicy(QSizePolicy.Policy.Fixed,QSizePolicy.Policy.Fixed)
                self.PPtx_Text_Path_label.setFrameShape(QFrame.Box)
                self.PPtx_Text_Path_label.setLineWidth(1)
                self.PPtx_Text_Path_label.setFont(QFont("Arial",15))
                self.PPtx_Text_Path_label.setWordWrap(True)
                vbox.addWidget(self.PPtx_Text_Path_label,1)
                

                self.PPtx_Text_label = QLabel(PPtx_Text,self)
                
                self.PPtx_Text_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                self.PPtx_Text_label.setMinimumSize(QSize(1000,50))
                self.PPtx_Text_label.setMaximumSize(QSize(1000,200))
                self.PPtx_Text_label.setSizePolicy(QSizePolicy.Policy.Fixed,QSizePolicy.Policy.Fixed)
                self.PPtx_Text_label.setFrameShape(QFrame.Box)
                self.PPtx_Text_label.setLineWidth(1)
                self.PPtx_Text_label.setWordWrap(True)
                self.PPtx_Text_label.setFont(QFont("Arial",12))
                self.PPtx_Text_label.setAlignment(Qt.AlignmentFlag.AlignVCenter)
                vbox.addWidget(self.PPtx_Text_label,2)
                
                self.ContentDisplay.append(widget)
                self.labelsLayout.addWidget(widget,y)

                y= y+1
                
                spacer = QSpacerItem(0,50)
                self.labelsLayout.addItem(spacer)
                self.ContentDisplay_spacer.append(spacer)
                y= y+1
            for i in range(0,2):
                ExtraWidget = QWidget()
                ExtraLayout = QVBoxLayout()
                ExtraWidget.setLayout(ExtraLayout)
                self.ContentDisplay.append(ExtraWidget)
                self.labelsLayout.addWidget(ExtraWidget,y)
                y= y+1

            #self.PPtx_Text_Path_label.setText(str(self.TextFrames))

            

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
        if int(page.text()) >= self.PPTX_PAGE:
            page.hide()
        else: 
            page.show()