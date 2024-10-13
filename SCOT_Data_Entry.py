#from fbs_runtime.application_context.PyQt5 import ApplicationContext  # used for fbs only
'''
Input files
    Company ID & Websites: f'~/Desktop/Scot_data_beta.xlsx'
    Elements Sorted & Prioritised: f'~/Desktop/Periodic_table.xlsx'
    Country slider mapping: pycountry.countries
Output
    SCOT_RD.xlsx
Extra Resources: Function'save_data' imports: from pox.shutils import find
'''
import sys
import pycountry
from PyQt5.QtWidgets import (QApplication,QFormLayout,QLineEdit,QVBoxLayout,QWidget,QDateEdit,
                             QHBoxLayout,QGridLayout,QPushButton,QListWidget,QLabel,QGroupBox,QRadioButton,
                             QSlider,QSizePolicy,QDesktopWidget,QMainWindow,QMessageBox)
#from pox.shutils import find
#from datetime import datetime
#import os
import numpy
import pandas as pd
from PyQt5 import QtCore
from PyQt5.QtCore import Qt,QTimer,QTime,QDateTime,QRegExp,QDate
from PyQt5.QtGui import QRegExpValidator,QDoubleValidator

class Window(QWidget):   #  need to redesign to use  QMainWindow

    def __init__(self):
        super().__init__()
        self.setWindowTitle("SCOT Data")
        self.setObjectName('mainwindow')

        # return current screen resolution
        self.screen_size = QDesktopWidget().screenGeometry(0)  # returns QRect, 0 =         main screen
        # set maximum geometry of window
        self.screen_height = int(self.screen_size.height()*.30)
        self.screen_width = int(self.screen_size.width()*.35)
        x_pos = 300
        y_pos = 300
        self.setGeometry(x_pos, y_pos, self.screen_width, self.screen_height)
        #self.resize(500, 350)
        # Initialize variables
        self.company_df=pd.read_excel(f'~/Desktop/SCOT_MT/Scot_data_beta.xlsx')[['name','conID','Website']]
        self.new_data_df=pd.DataFrame()
        self.company=''
        self.metal=''
        self.metal_string=''
        self.mine_local=QLabel('Antarctica') #initialize slider value to a non-mining district
        self.mine_local.setStyleSheet("font: 16pt \"Cambria\";\n""color: yellow;")

        regexp_a = QRegExp(r'^[a-zA-Z\s]*$')
        regexp_1= QRegExp(r'[0-9]+')
        self.numvalidator=QRegExpValidator(regexp_1)
        self.txtvalidator=QRegExpValidator(regexp_a)        
        #self.setCentralWidget(self.top_mid_layout)     not working tried APP. also
        self.ui_setup()

    def ui_setup(self):
        # Create an outer layout for main window
        outerLayout = QVBoxLayout()  
        # set up date time clock
        self.todays_date= QDateTime.currentDateTime().date().toString()

        # Create a GRID layout for the Titles (moved up from self.title=QLabel())
        self.banner_box_holder=QVBoxLayout()
        self.banner_layout=QGridLayout()
        
        # create inner box to put 'background' then add labels to grid layout *** ONLY GROUPBOX AND QMAINWINDOW CAN PUT JPG IN BORDER AND STRETCH
        self.box=QGroupBox(self)
        self.box.setObjectName("bannergroupbox")
        self.box.setFocusPolicy(Qt.NoFocus)
        self.box.setGeometry(10,10,int(self.screen_width*0.98),140)   #core-samples.jpg goldvein   core_pic2.jpg hammerrock Coregold
        stylesheet = '''#bannergroupbox{
                    border-image: url(/Users/EPIC/Desktop/SCOT_MT/greycore.jpg) 0 0 0 0 stretch stretch;
                    background-repeat: no-repeat;
                    }
                    '''
        self.box.setStyleSheet(stylesheet)
        # add titles to inner banner layout
        self.spacer_label=QLabel(self)
        self.spacer_label.setText('')
        self.banner_layout.addWidget(self.spacer_label,0,0,1,1)
        # add date time label
        self.datetime_label=QLabel(self)
        self.datetime_label.setObjectName("datetime")
        self.datetime_label.setGeometry(QtCore.QRect(30, 30, 150, 20))   #below title is about (270, 20, 50, 20))
        self.datetime_label.setStyleSheet("font: 16pt \"Cambria\";\n""color: black;")
        self.datetime_label.setAlignment(Qt.AlignRight)
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
        #date_time settext is called in show_time funnction call
        self.banner_layout.addWidget(self.datetime_label, 0, 2)
        #self.banner_layout.addWidget(self.time_label, 1, 0)
        
        self.title = QLabel(self)
        self.title.setAcceptDrops(False)
        #self.title.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.title.setObjectName("title")
        self.title.setText("SCOT Resource Data Entry")   
        self.title.setStyleSheet( "font: 30pt \"Cambria\";\n""color: black;""background-color: none;""border: none;")
        self.title.setFixedHeight(120)
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.title.setAlignment(QtCore.Qt.AlignTop)
        self.banner_layout.addWidget(self.title, 0, 1, 1,1)
        self.box.setLayout(self.banner_layout)
        self.banner_box_holder.addWidget(self.box)

        # Headers for input windows
        self.topLayout=QGridLayout()
        self.co_label=QLabel("Company Options")
        self.co_label.setStyleSheet( "font: 18pt \"Cambria\";\n""color: yellow;""background-color: blue;""border: 1px solid black;border-radius: 10px;")
        self.co_label.setFixedHeight(25)
        self.topLayout.addWidget(self.co_label, 0, 0, 1, 1)    # first version: QPushButton("Company Options")
        # column 1 will be used after company selection is made
        self.res_st_label=QLabel("Resource Options")
        self.res_st_label.setStyleSheet( "font: 18pt \"Cambria\";\n""color: yellow;""background-color: blue;""border: 1px solid black;border-radius: 10px;")
        self.co_label.setFixedHeight(25)
        self.topLayout.addWidget(self.res_st_label, 0, 3, 1, 1)     #QPushButton("Resource Options")
        
        # create company selection box as list_widget
        self.company_listWidget = QListWidget(self)
        self.company_listWidget.setObjectName("company list")
        self.company_listWidget.setFocusPolicy(Qt.ClickFocus)
        self.company_listWidget.setMaximumHeight(85)
        self.company_listWidget.setSelectionMode(1)                            
        self.company_listWidget.setSelectionRectVisible(True)                  
        self.company_listWidget.setAlternatingRowColors(True)                  
        self.company_listWidget.setGeometry(QtCore.QRect(50, 120, 100, 50))        
        # adding list of items to company list widget from the database of companies with finacial data
        list1=self.company_df['name'].tolist()
        list1=list(set(list1))
        list1.sort()
        co_list = list1
        self.company_listWidget.addItems(co_list)                             
        # prepare label for company clicked event
        self.label_4 = QLabel(self)
        self.co2_label=QLabel(self)
        self.label_4.setAlignment(QtCore.Qt.AlignLeft)
        #self.label_4.setGeometry(QtCore.QRect(390, 180, 400, 20))
        self.label_4.setStyleSheet("font: 18pt \"Cambria\";\n""color: blue;")
        self.label_4.setObjectName("label_4")
        self.company_listWidget.clicked.connect(self.company_clicked)          
        
        # Company choice accept Button
        self.co_choice_button = QPushButton(self)
        self.co_choice_button.setGeometry(QtCore.QRect(390, 120, 145, 60))
        self.co_choice_button.setStyleSheet("font: 20pt \"Cambria\";\n""color: black;\n""border: 2px solid black;""background-color: silver;border-radius: 10px;")
        self.co_choice_button.setStyleSheet("QPushButton::hover""{""color:yellow;""}""QPushButton::pressed""{""background-color:blue;border-radius: 10px;""}")
        self.co_choice_button.setObjectName("co_choice_button")
        self.co_choice_button.setText("1) Select Company")
        self.co_choice_button.clicked.connect(self.company_button_clicked)
        
        # create resource selector list widget
        self.metal_label=QLabel(self)
        self.metal_label.setStyleSheet("font: 18pt \"Cambria\";\n""color: yellow;")
        self.metal_label.setObjectName("metal_label")        
                
        self.res_listWidget =QListWidget(self)
        self.res_listWidget.setFocusPolicy(Qt.ClickFocus)
        self.res_listWidget.setMaximumHeight(85)
        self.res_listWidget.setSelectionMode(1) 
        self.res_listWidget.setSelectionRectVisible(True)   
        self.res_listWidget.setAlternatingRowColors(True)
        self.res_listWidget.setGeometry(QtCore.QRect(50, 320, 100, 50))
        self.res_listWidget.clicked.connect(self.res_clicked)
        self.res_listWidget.setObjectName("res_listWidget")
        # adding list of items to resource list widget The periodic table list has a 'priority key' for metals of interest
        #===>need to turn the tuple into a string
        all_metals=pd.read_excel(f'~/Desktop/SCOT_MT/Periodic_table.xlsx')
        all_metals.dropna(subset = ["Priority"], inplace=True)
        all_metals.sort_values(by=['Priority'],inplace=True)
        res_list=all_metals[['Element','Symbol']].apply(tuple,axis=1).tolist()
        joiner=" ".join
        res_list2=[joiner(resources) for resources in res_list]  #converts the tuples list into a list of strings        
        self.res_listWidget.addItems(res_list2)

        # Resource choice accept Button
        self.res_choice_button = QPushButton(self)
        self.res_choice_button.setGeometry(QtCore.QRect(390, 320, 145, 60))
        self.res_choice_button.setStyleSheet("font: 20pt \"Cambria\";\n""color: black;\n""border: 2px solid black;""background-color: silver;border-radius: 10px;")
        self.res_choice_button.setStyleSheet("QPushButton::hover""{""color:yellow;""}""QPushButton::pressed""{""background-color:blue;border-radius: 10px;""}")
        self.res_choice_button.setObjectName("res_choice_button")
        self.res_choice_button.setText("2) Select Metal") 
        self.res_choice_button.clicked.connect(self.res_button_clicked)
                
        # ADD COMPANY AND RESOURCES CHOICE BOXES TO LAYOUT
        self.topLayout.addWidget(self.company_listWidget,1,0,1,1)
        self.topLayout.addWidget(self.co_choice_button,3,0,1,1)        
        self.topLayout.addWidget(self.res_listWidget,1,3,1,1)
        self.topLayout.addWidget(self.res_choice_button,3,3,1,1)
        
        # Create Basic Mine inputs form :TOP-Mid Layout GRID Layout
        self.top_mid_layout=QGridLayout()
        self.create_data_forms()        
        #self.setCentralWidget(self.source_date)         #=====>>> DOESNT WORK
               
        # Create Metal resources input lines form :Mid level FORM Layout
        self.midLayout = QFormLayout()
        self.midLayout.setFormAlignment(Qt.AlignLeft)
        self.create_resouce_inputs()
        
        # Create a layout for the bottom buttons
        bottomLayout = QHBoxLayout()
        # Add some event buttons to the layout
        self.new_metal_btn=QPushButton("Save && Add New Metal")
        self.new_metal_btn.setStyleSheet("QPushButton::hover""{""color:yellow;""}""QPushButton::pressed""{""background-color:blue;border-radius: 10px;""}")
        self.new_mine_btn=QPushButton("Save && Add New Mine")
        self.new_mine_btn.setStyleSheet("QPushButton::hover""{""color:yellow;""}""QPushButton::pressed""{""background-color:blue;border-radius: 10px;""}")
        self.new_co_btn=QPushButton("Save && Change Company")
        self.new_co_btn.setStyleSheet("QPushButton::hover""{""color:yellow;""}""QPushButton::pressed""{""background-color:blue;border-radius: 10px;""}")               
        self.quit_btn=QPushButton("Export Data && Quit")  
        self.quit_btn.setStyleSheet("QPushButton::hover""{""color:yellow;""}""QPushButton::pressed""{""background-color:blue;border-radius: 10px;""}")
        #Lower layout button action functions
        self.new_metal_btn.clicked.connect(self.new_metal)
        self.new_mine_btn.clicked.connect(self.new_mine)
        self.new_co_btn.clicked.connect(self.new_company)        
        self.quit_btn.clicked.connect(self.exit_window)        
        
        bottomLayout.addWidget(self.new_metal_btn)
        bottomLayout.addWidget(self.new_mine_btn)
        bottomLayout.addWidget(self.new_co_btn)        
        bottomLayout.addWidget(self.quit_btn)
                
        # Nest the inner layouts into the outer layout
        # QVbox adds layouts vertically in the order you add them below
        outerLayout.addLayout(self.banner_box_holder)
        outerLayout.addLayout(self.topLayout)   
        outerLayout.addLayout(self.top_mid_layout)
        outerLayout.addLayout(self.midLayout)
        outerLayout.addLayout(bottomLayout)
        # Set the window's main layout
        self.setLayout(outerLayout)
        
    def res_button_focusout(self):
        self.source_date.setFocus(True)
        self.source_date.repaint()
        
    def showTime(self):
        current_time = QTime.currentTime()
        label_time = current_time.toString('hh:mm:ss')
        self.datetime_label.setText(f'{self.todays_date}\n{label_time}')

    def company_clicked(self):
        item=self.company_listWidget.currentItem().text()
        self.company=str(item)
        
    def company_button_clicked(self):
        #Echo company selection in a label followed by its website
        self.label_4.setStyleSheet("font: 16pt \"Cambria\";\n""color: yellow;")
        self.label_4.setText('You selected: '+ self.company)        
        #add webpage link
        message1= 'Click to go to Company Website'
        #get website from self company df if one exists
        try:
            website=self.company_df.loc[self.company_df['name']==self.company,'Website'].iloc[0]
            # for testing             website='https://www.aexgold.com/'
            #add double quotes to website string or it is not accepted by openLink function
            dq_website=f'"{website}"'
            linkTemplate='<a href={0}><font size=4 >{1}</a>'        
            self.label_4b = QLabel(self)
            self.label_4b.setText(linkTemplate.format(dq_website,message1))  
            self.label_4b.setOpenExternalLinks(True)
            self.label_4b.setObjectName("label_4b")
            self.label_4b.setAlignment(QtCore.Qt.AlignLeft)
            self.topLayout.addWidget(self.label_4b, 1, 2,1,1)
        except: pass
        self.topLayout.addWidget(self.label_4, 0, 2, 1,1)
        self.co2_label.setText(self.company+' Data Entry (<tab> to move, <enter/return> to validate input):')
        self.co2_label.setStyleSheet("font: 18pt \"Cambria\";\n""color: yellow;")
        self.topLayout.addWidget(self.co2_label,4,0,1,3)
        self.source_date.setFocus(True)
        
    def res_clicked(self):
        #self.metal=str('')
        item=self.res_listWidget.currentItem().text()
        self.metal=str(item)
        self.metal_label.setText('')
        
    def res_button_clicked(self):
        #last label before resource entry form on new layout 
        #self.metal_label=QLabel('Input Data For: '+ self.metal)
        self.metal_label.setStyleSheet("font: 18pt \"Cambria\";\n""color: yellow;")
        self.metal_label.setText('Input Data For: '+ self.metal)
        self.top_mid_layout.addWidget(self.metal_label,7,0,1,1)
        
        # LineEdit Focus fails to highlight Source Date QLE after resource button Clicked
        self.source_date.raise_()
        self.source_date.activateWindow()
        self.showNormal()
        self.source_date.setFocus(True)
        self.source_date.repaint()
        
    def onPressed(self):                  #enables move focus to next input widget after enter is pressed
        self.thiswidget=self.focusWidget()
        self.nextwidget=self.focusWidget().nextInFocusChain()
        #print('This one:',self.thiswidget)
        #print('Next one:',self.nextwidget)
        if self.thiswidget.objectName()=='production' or  self.thiswidget.objectName()=='sourcegroupbox' or self.thiswidget.objectName()=='source_date':
            self.source_date.setFocus()
            #print('on_pressed IF was TRUE',self.thiswidget.objectName())
        else:                                    
            self.thiswidget.focusNextChild()
            #print('on_pressed IF was False')
        
    def create_data_forms(self): 
        #create labels and input objects
        self.source_label=QLabel('Data Source:')        
        '''GET DATA SOURCE:   '''
        self.create_source_options()  #creates a group of radio buttons can choose only one
        
        self.source_date_label=QLabel('Source Date:(yyyy-mm-dd):<TAB>fwd')
        self.source_date=QDateEdit(date=QDate.currentDate(), calendarPopup=True,objectName='source_date')
        self.source_date.setDisplayFormat("yyyy-MMM-dd")
        self.source_date.setDateRange(QDate(1990, 1, 1), QDate.currentDate())
        self.source_date.editingFinished.connect(lambda:self.onPressed)  #   .editingFinished.   (lambda:) OR .dateChanged.
  
        #self.source_date=QLineEdit()
        self.source_date.setFixedWidth(140)
        self.source_date.setAlignment(Qt.AlignRight)
        self.source_date.editingFinished.connect(lambda: self.onPressed)
        #self.source_date.setInputMask('9999-99-99')
        #self.source_date.returnPressed.connect(self.onPressed)
        self.mine_name_label=QLabel('Mine/Site Name:')
        self.mine_name=QLineEdit()
        self.mine_name.setValidator(self.txtvalidator) 
        self.mine_name.setAlignment(Qt.AlignLeft)
        self.mine_name.returnPressed.connect(self.onPressed)        
        #self.mine_name.setInputMask('aaaaaaaaaaaaaaaaaaaaaaaa')
        self.mine_acres_label=QLabel('Hectares on Site:')         
        self.mine_acres=QLineEdit()
        self.mine_acres.setValidator(self.numvalidator)
        self.mine_acres.returnPressed.connect(self.onPressed) 
        #self.mine_acres.setInputMask('00000000')
        self.holes_drilled_label=QLabel('Holes Drilled:')
        self.holes_drilled=QLineEdit()
        self.holes_drilled.setValidator(self.numvalidator)
        self.holes_drilled.returnPressed.connect(self.onPressed) 
        #self.holes_drilled.setInputMask('000000')                           
        self.meters_drilled_label=QLabel('Meters Drilled:')
        self.meters_drilled=QLineEdit()
        self.meters_drilled.setObjectName('meters_drilled') 
        self.meters_drilled.setValidator(self.numvalidator)        
        self.meters_drilled.returnPressed.connect(self.onPressed) 
        #self.meters_drilled.setInputMask('0000000')

        ''' Get Mine Status'''
        self.status_label=QLabel('Mine Status:')
        self.create_status_options()
        '''GET MINE LOCATION'''        
        self.mine_local_label=QLabel('Mine Location: (slide to select ==>)')
        self.create_country_slider()         # dont want to create a new slider if there is already a name for Mine

        # Position labels and objects        
        self.top_mid_layout.addWidget(self.source_label,1,0,1,1)
        self.top_mid_layout.addWidget(self.sourcegroupBox,1,1,1,3)    
        self.top_mid_layout.addWidget(self.source_date_label,2,0,1,1)
        self.top_mid_layout.addWidget(self.source_date,2,1,1,1)
        self.top_mid_layout.addWidget(self.mine_name_label,3,0,1,1)
        self.top_mid_layout.addWidget(self.mine_name,3,1,1,1)
        self.top_mid_layout.addWidget(self.mine_acres_label,3,2,1,1)
        self.top_mid_layout.addWidget(self.mine_acres,3,3,1,1)
        self.top_mid_layout.addWidget(self.holes_drilled_label,4,0,1,1)
        self.top_mid_layout.addWidget(self.holes_drilled,4,1,1,1)
        self.top_mid_layout.addWidget(self.meters_drilled_label,4,2,1,1)                                      
        self.top_mid_layout.addWidget(self.meters_drilled,4,3,1,1) 
        # Mine Status groupbox radio buttons
        self.top_mid_layout.addWidget(self.status_label,5,0,1,1)
        self.top_mid_layout.addWidget(self.statusgroupBox,5,1,1,3)
        # Mine Location Slider
        self.top_mid_layout.addWidget(self.mine_local_label,6,0,1,1) 
        self.top_mid_layout.addWidget(self.slider_grid,6,1,2,3)      
        
    def create_resouce_inputs(self):
        #create lower mid_level QEditlines with col headers from data base
        #'tonnage_mt','ave_grade','proven','probable','measured','indicated','inferred','pay_depth','npv10','production'
        self.tonnage_mt=QLineEdit() 
        self.tonnage_mt.setValidator(self.numvalidator)
        self.tonnage_mt.returnPressed.connect(self.onPressed)
        #self.tonnage_mt.setInputMask('00000000')
        self.midLayout.addRow("Gross Rock Tonnage mt:",self.tonnage_mt)
        
        self.ave_grade=QLineEdit()
        self.ave_grade.setValidator(QDoubleValidator(0.99,99.99,2))
        self.ave_grade.returnPressed.connect(self.onPressed)
        #self.ave_grade.setInputMask('00.00')
        self.midLayout.addRow("Ave Grade (xx.xx g/t):", self.ave_grade)
        
        self.proven=QLineEdit()
        self.proven.setValidator(self.numvalidator)
        self.proven.returnPressed.connect(self.onPressed)
        #self.proven.setInputMask('00000000')
        self.midLayout.addRow("Proven:", self.proven)
        
        self.probable=QLineEdit()
        self.probable.setValidator(self.numvalidator)
        self.probable.returnPressed.connect(self.onPressed)
        #self.probable.setInputMask('00000000')
        self.midLayout.addRow("Probable:", self.probable)
        
        self.measured=QLineEdit()
        self.measured.setValidator(self.numvalidator)
        self.measured.returnPressed.connect(self.onPressed)
        #self.measured.setInputMask('00000000')
        self.midLayout.addRow("Measured:", self.measured)

        self.indicated=QLineEdit()
        self.indicated.setValidator(self.numvalidator)
        self.indicated.returnPressed.connect(self.onPressed)
        #self.indicated.setInputMask('00000000')
        self.midLayout.addRow("Indicated:", self.indicated)

        self.inferred=QLineEdit()
        self.inferred.setValidator(self.numvalidator)
        self.inferred.returnPressed.connect(self.onPressed)
        #self.inferred.setInputMask('00000000')
        self.midLayout.addRow("Inferred:", self.inferred)

        self.pay_depth=QLineEdit()
        self.pay_depth.setValidator(self.numvalidator)
        self.pay_depth.returnPressed.connect(self.onPressed)
        #self.pay_depth.setInputMask('00000')
        self.midLayout.addRow("PayDepth (m)", self.pay_depth)

        self.npv10=QLineEdit()
        self.npv10.setValidator(self.numvalidator)
        self.npv10.returnPressed.connect(self.onPressed) 
        #self.npv10.setInputMask('$000000000')
        self.midLayout.addRow("NPV10: ($)", self.npv10)

        self.production=QLineEdit(objectName='production')
        self.production.setValidator(self.numvalidator)
        #self.production.setInputMask('0000000')
        self.production.returnPressed.connect(self.onPressed)        
        self.midLayout.addRow("Production (oz|tn/lbs):", self.production)
        
    def create_source_options(self):
        self.sourcegroupBox = QGroupBox()
        self.sourcegroupBox.setObjectName("sourcegroupbox")
        self.sourcegroupBox.setFocusPolicy(Qt.ClickFocus)         #=======>>>>
        self.source1 = QRadioButton("Corp_Presentation")
        self.source1.toggled.connect(self.on_selected)
        self.source2 = QRadioButton("Annual_Report")
        self.source2.toggled.connect(self.on_selected)
        self.source3 = QRadioButton("Geologic_Report")
        self.source3.toggled.connect(self.on_selected)
        self.source1.setChecked(True)                       # if you deselect here you need error chacking in capture data function

        self.hbox = QHBoxLayout()
        self.hbox.addWidget(self.source1)
        self.hbox.addWidget(self.source2)
        self.hbox.addWidget(self.source3)
        self.hbox.addStretch(1)
        self.sourcegroupBox.setLayout(self.hbox)
        
    def on_selected(self):
        rad_but_choice=self.sender()
        if rad_but_choice.isChecked():
            self.source=rad_but_choice.text()
        #print(rad_but_choice.text())
        
    def create_status_options(self):
        self.statusgroupBox = QGroupBox()
        self.statusgroupBox.setObjectName("statusgroupbox")
        self.status1 = QRadioButton("Exploration")
        self.status1.toggled.connect(self.status_selected)
        self.status2 = QRadioButton("Development")
        self.status2.toggled.connect(self.status_selected)
        self.status3 = QRadioButton("Production")
        self.status3.toggled.connect(self.status_selected)
        self.status1.setChecked(True)     # if you deselect here you need error chacking in capture data function

        self.hbox2 = QHBoxLayout()
        self.hbox2.addWidget(self.status1)
        self.hbox2.addWidget(self.status2)
        self.hbox2.addWidget(self.status3)
        self.hbox2.addStretch(1)
        self.statusgroupBox.setLayout(self.hbox2)

    def status_selected(self):
        status_choice=self.sender()
        if status_choice.isChecked():
            self.status=status_choice.text()
        
    def create_country_slider(self):
        self.slider_grid=QGroupBox()
        self.slider_grid.setObjectName("slidergroupbox")
        self.country_slider=QSlider()
        self.country_slider.setOrientation(Qt.Horizontal)
        self.country_slider.setTickPosition(QSlider.TicksAbove)
        self.country_slider.setTickInterval(1)
        self.country_slider.setMinimum(0)
        self.country_slider.setMaximum(len(pycountry.countries)-1)  #slider index starts a zero
        self.country_slider.valueChanged.connect(self.changed_slider)

        self.vbox2=QVBoxLayout()        
        self.vbox2.addWidget(self.country_slider)
        self.vbox2.addWidget(self.mine_local)   
        self.slider_grid.setLayout(self.vbox2) 
        
    def changed_slider(self):
        value = self.country_slider.value()
        countries = sorted([country.name for country in pycountry.countries] , key=lambda x:x)
        self.mine_local.setText(str(countries[value]))

    def new_company(self):
        #save data first then clear Layout capture data is done in call to new_metal 
        self.capture_data()
        
        self.company=''
        self.label_4.setText('You selected: nothing !')
        self.label_4b.setText('                      ')
        self.co2_label.setText('')

        self.new_mine()
      
    def reset_tab_order(self):
        self.setTabOrder(self.source_date.focusProxy(), self.mine_name.focusProxy())
        self.setTabOrder(self.mine_name.focusProxy(), self.mine_acres.focusProxy())
        self.setTabOrder(self.mine_acres.focusProxy(), self.holes_drilled.focusProxy())
        self.setTabOrder(self.holes_drilled.focusProxy(), self.meters_drilled.focusProxy())
        self.setTabOrder(self.meters_drilled.focusProxy(), self.tonnage_mt.focusProxy())
        self.setTabOrder(self.tonnage_mt.focusProxy(), self.ave_grade.focusProxy())
        self.setTabOrder(self.ave_grade.focusProxy(), self.proven.focusProxy())
        self.setTabOrder(self.proven.focusProxy(), self.probable.focusProxy())
        self.setTabOrder(self.probable.focusProxy(), self.measured.focusProxy())
        self.setTabOrder(self.measured.focusProxy(), self.indicated.focusProxy())
        self.setTabOrder(self.indicated.focusProxy(), self.inferred.focusProxy())
        self.setTabOrder(self.inferred.focusProxy(), self.pay_depth.focusProxy())
        self.setTabOrder(self.pay_depth.focusProxy(), self.npv10.focusProxy())
        self.setTabOrder(self.npv10.focusProxy(), self.production.focusProxy())
        self.setTabOrder(self.production.focusProxy(), self.source_date.focusProxy())
        
    def new_mine(self):
        #save data first then clear Layout capture data is done in call to new_metal
        self.new_metal() 
        
        self.clearLayout(self.top_mid_layout) 
        self.mine_local=QLabel('Antarctica') 
        
        self.create_data_forms()                #should add blank forms ONLY IF ITS DELETED IN clearLayout
        self.source_date.setFocus(True)    
                
    def new_metal(self):
        #save data first then clear Layout
        self.capture_data()
        
        self.clearLayout(self.midLayout)  
        self.metal_label = QLabel(self)
        self.metal_label.setText('')       #('Input Data For: '+ self.metal)
        self.create_resouce_inputs()       #should add blank form  ONLY IF ITS DELETED IN clearLayout
        self.tonnage_mt.setFocus(True)
               
    def clearLayout(self,layout):           #call for each layout as required do layout refresh in call
        self.metal=''
        self.metal_label.setText('')
        while layout.count():
            child = layout.takeAt(0)
            childWidget = child.widget()
            if childWidget:
                childWidget.setParent(None)
                childWidget.deleteLater()
                
    def capture_data(self):
        # adds a row of data to the new dataframe for this session
        # checks that company and metal have been selected and at least 1 of proven or measured is entered before adding data
        if len(self.company)>0 and len(self.metal)>1 and (len(self.proven.text())>0 or len(self.measured.text())>0):
            website=self.company_df.loc[self.company_df['name']==self.company,'Website'].iloc[0]
            conID=int(self.company_df.loc[self.company_df['name']==self.company,'conID'].iloc[0])                     
            #print(self.company)
            data_dict={'name':self.company,        
                       'website':website,
                       'conID':conID,
                       'source':self.source,
                       'source_date':self.source_date.text(),
                       'mine_name':self.mine_name.text(),
                       'mine_acres':self.mine_acres.text(),
                       'holes_drilled':self.holes_drilled.text(),
                       'meters_drilled':self.meters_drilled.text(),
                       'status':self.status,
                       'mine_local':self.mine_local.text(),
                       'mine_metal':self.metal,
                       'tonnage_mt':self.tonnage_mt.text(),
                       'ave_grade':self.ave_grade.text(),
                       'proven':self.proven.text(),
                       'probable':self.probable.text(),
                       'measured':self.measured.text(),
                       'indicated':self.indicated.text(),
                       'inferred':self.inferred.text(),
                       'pay_depth':self.pay_depth.text(),
                       'npv10':self.npv10.text(),
                       'production':self.production.text()
                      }
            self.new_data_df=self.new_data_df.append(data_dict,ignore_index=True)

    def save_data(self):
        import os
        from pox.shutils import find
        from datetime import datetime
        
        #will export the captured data from the new session dataframe to excel (appending where appropriate)
        #print(self.new_data_df)
        timestamp = datetime.now().strftime('-%Y-%b-%d-%HH-%MM')
        
        #find file on local directory
        source_path=find('SCOT_RD.xlsx')      # uses pox.shutils, returns a list of strings index [0] is first instance

        #if file found append new data otherwise create new file
        if len(source_path)>0:
            original_df=pd.read_excel(source_path[0])        # use the first instance of the file in the list
            original_source=source_path[0]                   # save this string path to re-write appended df
    
            #create backup and rename the old file but add time stamp to the name
            current_file=os.path.splitext(source_path[0])[0]  #source is a list take first string element and remove the ext with splitext [foo, .ext]
            back_up_file=(f'{current_file}{timestamp}.xlsx')  #/Users/EPIC/Desktop/core_pic2-2022-Mar-30-17H-27M.xlsx
            os.rename(source_path[0],back_up_file)              # Left into Right
            #append data to original
            appended_df=original_df.append(self.new_data_df,ignore_index=True)
            appended_df.to_excel(original_source,index=False)
            exit_path = source_path[0]
            #print(exit_path)
        else: 
            newpath=os.getcwd()                               #write a new file into the current working dir
            newfile='SCOT_RD.xlsx'
            self.new_data_df.to_excel(f'{newpath}/{newfile}',index=False)
            exit_path=f'{newpath}/{newfile}'
            
        # show confirmation that data was saved        
        messageBox = QMessageBox()
        messageBox.setText(f'Data Saved at {exit_path}\n\n Have Great Day!!')
        messageBox.setStandardButtons(QMessageBox.Ok)   #has other options
        messageBox.exec()
        #self.close() 
    
    def exit_window(self):
        self.capture_data()
        if len(self.new_data_df.index)>0:
            self.save_data()
        self.deleteLater()
        self.close()
        self.destroy()
        sys.exit()
        #os._exit(0)   #kills the kernel

def main():
    #appctxt = ApplicationContext()               #for fbs only
    if not QtCore.QCoreApplication.instance():
        App = QApplication(sys.argv)
    else:
        App = QApplication.instance()
    # create the instance of our Window
    window = Window()
    window.show()
    #exit_code = appctxt.app.exec()              #for fbs only 
    App.exec()
    
    
if __name__ == '__main__':
   main()
