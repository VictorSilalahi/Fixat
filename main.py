import sys
import time
import datetime
from inc import connection

import xlsxwriter
import win32com.client
import os

from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtWidgets

from fpdf import FPDF

from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt


from inc import frmAddEdit

class MainWin(QtWidgets.QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWin()

    def setWin(self):
        self.setWindowTitle("FIXAT")
        # layout
        self.winLayout = QtWidgets.QGridLayout(self)

        # UI components
        gBoxLeft = QtWidgets.QGroupBox("Choose Type:")
        self.tvwDatType =  QtWidgets.QTreeWidget()

        self.twCat = QtWidgets.QTreeWidgetItem(["By Category"])
        self.fillByType("By Category")
        self.tvwDatType.insertTopLevelItem(0,self.twCat)
        
        self.twLoc = QtWidgets.QTreeWidgetItem(["By Location"])
        self.fillByType("By Location")
        self.tvwDatType.insertTopLevelItem(1,self.twLoc)
        self.tvwDatType.expandAll()
        self.tvwDatType.header().hide()

        vboxLeft = QtWidgets.QVBoxLayout()
        vboxLeft.addWidget(self.tvwDatType,3)
        
        # tests
        self.tabs = QtWidgets.QTabWidget()
        self.tabByCategory = QtWidgets.QWidget()
        self.tabByLocation = QtWidgets.QWidget()

        self.tabs.addTab(self.tabByCategory,"By Category")
        self.tabs.addTab(self.tabByLocation,"By Location")
        
        lOutCategory = QtWidgets.QGridLayout()
        self.figCategory = plt.figure(figsize=(1,1))
        self.canvCategory = FigureCanvas(self.figCategory)
        lOutCategory.addWidget(self.canvCategory)
        self.ax1 = self.figCategory.add_subplot(111)
        self.ax1.set_title("Berdasarkan Category")
        
        self.tabByCategory.setLayout(lOutCategory)

        lOutLocation = QtWidgets.QGridLayout()
        self.figLocation = plt.figure(figsize=(1,1))
        self.canvLocation = FigureCanvas(self.figLocation)
        lOutLocation.addWidget(self.canvLocation)
        self.ax2 = self.figLocation.add_subplot(111)
        self.ax2.set_title("Berdasarkan Lokasi")
        
        self.tabByLocation.setLayout(lOutLocation)
    

        vboxLeft.addWidget(self.tabs,3)
        
        vboxLeft.addStretch(1)
        gBoxLeft.setLayout(vboxLeft)
        
        self.winLayout.addWidget(gBoxLeft,0,0,15,2)

        gBoxRight = QtWidgets.QGroupBox("Asset List:")
        self.btnAdd = QtWidgets.QPushButton("Add")
        self.btnAdd.setIcon(QtGui.QIcon("icons/add.png"))
        self.btnAdd.clicked.connect(self.addItem)
        self.btnEdit = QtWidgets.QPushButton("Edit")
        self.btnEdit.setIcon(QtGui.QIcon("icons/edit.png"))
        self.btnEdit.clicked.connect(self.editItem)
        self.btnDel = QtWidgets.QPushButton("Delete")
        self.btnDel.setIcon(QtGui.QIcon("icons/del.png"))
        self.tblAsset = QtWidgets.QTableWidget(0,8)
        self.tblAsset.setSelectionMode( QtWidgets.QAbstractItemView.MultiSelection )
        self.tblAsset.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.tblAsset.setHorizontalHeaderLabels( ["#Asset No","#SN","Description","Acq Date","Acq Cost","Depr Method","Useful Life","Current Value"] )
        self.pBar = QtWidgets.QProgressBar()
        lblMedia = QtWidgets.QLabel("Media to print:")
        self.cmbPrintingMedia = QtWidgets.QComboBox()
        self.cmbPrintingMedia.addItem("XLS")
        self.cmbPrintingMedia.addItem("PDF")
        self.btnPrint = QtWidgets.QPushButton("Print")
        self.btnPrint.setIcon(QtGui.QIcon("icons/print.png"))
        
        vboxRight = QtWidgets.QGridLayout()
        vboxRight.addWidget(self.btnAdd,0,0)
        vboxRight.addWidget(self.btnEdit,0,1)
        vboxRight.addWidget(self.btnDel,0,2)
        #vboxRight.addWidget(self.pBar,1,0,1,2)
        vboxRight.addWidget(self.tblAsset,1,0,15,3)
        vboxRight.addWidget(self.cmbPrintingMedia,16,1)
        vboxRight.addWidget(lblMedia,16,0)
        lblMedia.setAlignment(QtCore.Qt.AlignRight)
        vboxRight.addWidget(self.btnPrint,16,2)
        
        gBoxRight.setLayout(vboxRight)
        
        self.winLayout.addWidget(gBoxRight,0,3,15,5)
        
        # popup menu
        self.tvwDatType.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.tvwDatType.customContextMenuRequested.connect(self.tvwDatRightClicked)
        
        # standard click event on tableWidget
        self.tvwDatType.clicked.connect(self.fillTable)
        
        # button click event
        self.btnDel.clicked.connect(self.delAsset)
        self.btnPrint.clicked.connect(self._print)
        # textbox for data editing in qtreewidget
        self.lnEdit = QtWidgets.QLineEdit()
        
        
        # show window
        self.showMaximized()
        
        # fill graph
        self.fillGraph()
        
    def tvwDatRightClicked(self,pos):
        #if self.tvwDatType.indexOfTopLevelItem(self.tvwDatType.currentItem())==-1:
        #    QtWidgets.QMessageBox.about( self,"Data", self.tvwDatType.currentItem().parent().text(0)+" - "+self.tvwDatType.currentItem().text(0) )
        #else:   
        #    return self.tvwDatType.currentItem(),-1            
        if self.tvwDatType.indexOfTopLevelItem(self.tvwDatType.currentItem())==-1:
            # create menu
            self.popupmenu = QtWidgets.QMenu()
            actionAdd = QtWidgets.QAction("Add")
            actionAdd.setIcon(QtGui.QIcon("icons/add.png"))
            actionEdit = QtWidgets.QAction("Edit")
            actionEdit.setIcon(QtGui.QIcon("icons/edit.png"))
            actionDel = QtWidgets.QAction("Delete")
            actionDel.setIcon(QtGui.QIcon("icons/del.png"))
            actionAdd.triggered.connect(self.addCriteria)
            actionEdit.triggered.connect(self.editCriteria)
            actionDel.triggered.connect(self.delCriteria)
            self.popupmenu.addAction(actionAdd)
            self.popupmenu.addAction(actionEdit)
            self.popupmenu.addAction(actionDel)
            action = self.popupmenu.exec_(self.tvwDatType.mapToGlobal(pos))
        
    def addCriteria(self):
        txtNew, okPressed = QtWidgets.QInputDialog.getText(self, "New Category/Location","New Data:", QtWidgets.QLineEdit.Normal, "")
        if okPressed and txtNew != '':
            cur = connection.connection()
            if (self.tvwDatType.currentItem().parent().text(0)=="By Category"):
                cur.execute("select count(*) as amount from tCategories where Name='"+txtNew+"'")
                rAmount = cur.fetchone()
                if rAmount[0]>0:
                    QtWidgets.QMessageBox.about( self,"New Category", "This new category already exist!" )
                else:
                    cur.execute("insert into tCategories(Name) values('"+txtNew+"')")
                    connection.con.commit()
                    self.twCat.addChild( QtWidgets.QTreeWidgetItem([txtNew]) )
            else:
                cur.execute("select count(*) as amount from tLocations where LocName='"+txtNew+"'")
                rAmount = cur.fetchone()
                if rAmount[0]>0:
                    QtWidgets.QMessageBox.about( self,"New Location", "This new location already exist!" )
                else:
                    cur.execute("insert into tLocations(LocName) values('"+txtNew+"')")
                    connection.con.commit()
                    self.twLoc.addChild( QtWidgets.QTreeWidgetItem([txtNew]) )

                
        else:
            QtWidgets.QMessageBox.about( self,"New Data", "Data can not empty!" )
    
    def editCriteria(self):
        #if (self.tvwDatType.currentItem().parent().text(0)=="By Category") or (self.tvwDatType.currentItem().parent().text(0)=="By Location" ):
        #self.tvwDatType.openPersistentEditor( self.tvwDatType.currentItem(),0)
        self.val = self.tvwDatType.currentItem().text(0)
        itm = self.tvwDatType.itemFromIndex(self.tvwDatType.selectedIndexes()[0])
        column = self.tvwDatType.currentColumn()        

        self.lnEdit.setText(self.val)
        self.tvwDatType.setItemWidget(itm,column,self.lnEdit)
        self.lnEdit.show()
        self.lnEdit.setFocus()
        
        
    def delCriteria(self):
        cur = connection.connection()
        if (self.tvwDatType.currentItem().parent().text(0)=="By Category"):
            cur.execute("select count(*) as amount from tCategories,tAssets where tCategories.CategoryID=tAssets.CategoryID and tCategories.Name='"+self.tvwDatType.currentItem().text(0)+"'")
            rAmount = cur.fetchone()
            if rAmount[0]>0:
                QtWidgets.QMessageBox.about( self,"Delete Category", "Can not delete this category. This category has been use on several items!" )
            else:
                cur.execute("delete from tCategories where Name='"+self.tvwDatType.currentItem().text(0)+"'" )
                connection.con.commit()
                self.tvwDatType.currentItem().parent().removeChild( self.tvwDatType.currentItem() )
        else:
            cur.execute("select count(*) as amount from tLocations,tAssets where tLocations.LocationID=tAssets.LocationID and tLocations.LocName='"+self.tvwDatType.currentItem().text(0)+"'")
            rAmount = cur.fetchone()
            if rAmount[0]>0:
                QtWidgets.QMessageBox.about( self,"Delete Location", "Can not delete this location. This location has been use on several items!" )
            else:
                cur.execute("delete from tLocations where LocName='"+self.tvwDatType.currentItem().text(0)+"'" )
                connection.con.commit()
                self.tvwDatType.currentItem().parent().removeChild( self.tvwDatType.currentItem() )
            
        
    def saveCriteria(self,item):
        print(self.tvwDatType.currentItem().text(0))
        
    def fillByType(self,vtype):
        cur = connection.connection()
        if vtype=="By Category":
            cur.execute("select name from tCategories")
            rCat = cur.fetchall()
            for row in rCat:
                self.twCat.addChild( QtWidgets.QTreeWidgetItem([row[0]]) )
        else:
            cur.execute("select locname from tLocations")
            rLoc = cur.fetchall()
            for row in rLoc:
                self.twLoc.addChild( QtWidgets.QTreeWidgetItem([row[0]]) )
            
    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape:
            self.lnEdit.clearFocus()
            self.lnEdit.hide()
            
    def fillTable(self):
        self.tblAsset.setRowCount(0)
        #cols = 9
        time.sleep(1)
        
        cur = connection.connection()
        con = connection.con
        if (self.tvwDatType.currentItem().parent().text(0)=="By Category"):
            cur.execute( "select tAssets.AssetNo, tAssets.SN, tAssets.AsDesc, tAssets.AcqDate, tCategories.Name, tLocations.LocName, tAssets.AcqCost, tAssets.DepMeth, tAssets.UsefulLive from tAssets, tCategories, tLocations where tAssets.CategoryID=tCategories.CategoryID and tAssets.LocationID=tLocations.LocationID and tCategories.Name='" + self.tvwDatType.currentItem().text(0) + "' order by tAssets.AssetNo")
        else:
            cur.execute( "select tAssets.AssetNo, tAssets.SN, tAssets.AsDesc, tAssets.AcqDate, tCategories.Name, tLocations.LocName, tAssets.AcqCost, tAssets.DepMeth, tAssets.UsefulLive from tAssets, tCategories, tLocations where tAssets.CategoryID=tCategories.CategoryID and tAssets.LocationID=tLocations.LocationID and tLocations.LocName='" + self.tvwDatType.currentItem().text(0) + "' order by tAssets.AssetNo")
        
        rows = cur.fetchall()
        self.tblAsset.setRowCount(len(rows))
        no=0
        for r in rows:
            self.tblAsset.setItem(no,0,QtWidgets.QTableWidgetItem(r[0]))
            self.tblAsset.setItem(no,1,QtWidgets.QTableWidgetItem(r[1]))
            self.tblAsset.setItem(no,2,QtWidgets.QTableWidgetItem(r[2]))
            self.tblAsset.setItem(no,3,QtWidgets.QTableWidgetItem(r[3]))
            self.tblAsset.setItem(no,4, QtWidgets.QTableWidgetItem( '{:0,.0f}'.format(r[6]) ) )
            self.tblAsset.item(no,4).setTextAlignment(QtCore.Qt.AlignRight)
            self.tblAsset.setItem(no,5,QtWidgets.QTableWidgetItem(r[7]))
            self.tblAsset.item(no,5).setTextAlignment(QtCore.Qt.AlignRight)
            self.tblAsset.setItem(no,6,QtWidgets.QTableWidgetItem( str(r[8])))
            self.tblAsset.item(no,6).setTextAlignment(QtCore.Qt.AlignRight)
            dNow = datetime.date.today()
            dAcq = datetime.datetime.strptime(r[3],"%m/%d/%Y")
            deltYears = dNow.year - dAcq.year
            cV = self.currVal(r[6], r[7], deltYears,r[8])
            self.tblAsset.setItem(no,7,QtWidgets.QTableWidgetItem( '{:0,.0f}'.format(cV) ) )
            self.tblAsset.item(no,7).setTextAlignment(QtCore.Qt.AlignRight)
            no=no+1
            
        header = self.tblAsset.horizontalHeader()
        header.setSectionResizeMode(0,QtWidgets.QHeaderView.ResizeToContents)
    
    def delAsset(self):
        listRow=[]
        rowSelected = 0
        for i in range(self.tblAsset.rowCount()):
            if self.tblAsset.item(i,0).isSelected()== True:
                rowSelected=rowSelected+1
                listRow.append(self.tblAsset.item(i,0).text())
        if rowSelected>0:
            ans = QtWidgets.QMessageBox.question( self, 'Delete Assets', "Do you want to delete some assets?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No )
            if ans == QtWidgets.QMessageBox.Yes:
                cur = connection.connection()
                con = connection.con
                for j in range(len(listRow)):
                    for k in range(self.tblAsset.rowCount()):
                        if self.tblAsset.item(k,0).text()==listRow[j]:
                            cur.execute( "delete from tAssets where AssetNo='"+listRow[j]+"'")
                            connection.con.commit()
                            print(listRow[j])
                            self.tblAsset.removeRow(k)
                            break
                        else:
                            continue
        else:
            QtWidgets.QMessageBox.about( self,"Delete Asset", "Please select one or more item in the asset table!" )
            
    def currVal(self,aC,dType,deltY,uL):
        lastVal=aC
        if dType=="SLN":
            if deltY>0:
                depVal = aC/uL
                for i in range(deltY):
                    lastVal = lastVal-depVal
        elif dType=="DDB":
            if deltY>0:
                depVal = 2/uL
                for i in range(deltY):
                    lastVal = lastVal - (lastVal*depVal)
        else:
            sumOfYears = 0
            if deltY>0:
                for i in range(1,uL):
                    sumOfYears = sumOfYears + i
                for j in range(deltY):
                    lastVal = lastVal - (lastVal*(i/sumOfYears))
        if lastVal<0:
            lastVal=0
        return lastVal
    
    def addItem(self):
        self.assetOp = "add"
        self.winAddEdit = frmAddEdit.winAdd(parent=self)
    
    def editItem(self):
        rowSelected = 0
        for i in range(self.tblAsset.rowCount()):
            if self.tblAsset.item(i,0).isSelected()== True:
                rowSelected=rowSelected+1
        if rowSelected>1 or rowSelected==0:
            QtWidgets.QMessageBox.about( self,"Edit Asset", "Please select one item in the asset table for editing operation!" )
        else:
            self.assetOp = "edit"
            self.winAddEdit = frmAddEdit.winAdd(parent=self)
    
    def _print(self):
        if self.tblAsset.rowCount()==0:
            QtWidgets.QMessageBox.about( self,"Asset Data", "There's no data!" )
        else:
            if self.cmbPrintingMedia.currentText()=="XLS":
                wbook = xlsxwriter.Workbook('data.xlsx')
                wsheet = wbook.add_worksheet("Data")
                
                merge_format =wbook.add_format({'bold': True, 'font_color': 'blue','font_size':18})
                wsheet.merge_range("B2:E3","Asset Inventory",merge_format)
                cell_format=wbook.add_format({'bold': True,'align':'right'})
                wsheet.write('B4', 'Company Name :',cell_format)
                wsheet.write('C4', 'Audyne LLC')
                d = datetime.datetime.now()
                wsheet.write('B5', 'Date :',cell_format)
                wsheet.write('C5', str(d.month)+"/"+str(d.day)+"/"+str(d.year))
                if self.tvwDatType.currentItem().parent().text(0)=="By Category":
                    wsheet.write('E4', 'Category : '+self.tvwDatType.currentItem().text(0) ,cell_format)
                else:
                    wsheet.write('E4', 'Location : '+self.tvwDatType.currentItem().text(0) ,cell_format)
                cell_format=wbook.add_format({ 'border':1, 'fg_color':'yellow','font_size':12 })
                wsheet.write('B6','#No',cell_format)
                wsheet.write('C6','#Item Number',cell_format)
                wsheet.write('D6','#Serial Number',cell_format)
                wsheet.write('E6','Description',cell_format)
                wsheet.write('F6','Acq Date',cell_format)
                wsheet.write('G6','Acq Cost',cell_format)
                wsheet.write('H6','Depr Method',cell_format)
                wsheet.write('I6','Usefull Live',cell_format)
                wsheet.write('J6','CurrentValue',cell_format)
                
                col=1
                for x in range(self.tblAsset.rowCount()):
                    if x==self.tblAsset.rowCount()-1:
                        cell_format=wbook.add_format({ 'left':1, 'right':1, 'num_format':'##,###,###,###','bottom':1 })
                    else:
                        cell_format=wbook.add_format({ 'left':1, 'right':1, 'num_format':'##,###,###,###' })
                    wsheet.write(x+6,col,int(x+1), cell_format)
                    wsheet.write(x+6,col+1,self.tblAsset.item(x,0).text(), cell_format)
                    wsheet.write(x+6,col+2,self.tblAsset.item(x,1).text(), cell_format)
                    wsheet.write(x+6,col+3,self.tblAsset.item(x,2).text(), cell_format)
                    wsheet.write(x+6,col+4,self.tblAsset.item(x,3).text(), cell_format)
                    wsheet.write(x+6,col+5,self.tblAsset.item(x,4).text() , cell_format)
                    wsheet.write(x+6,col+6,self.tblAsset.item(x,5).text(), cell_format)
                    wsheet.write(x+6,col+7,self.tblAsset.item(x,6).text(), cell_format)
                    wsheet.write(x+6,col+8,self.tblAsset.item(x,7).text(), cell_format)
                    
                
                wbook.close()            

                xl=win32com.client.Dispatch("Excel.Application")
                xl.Visible = True
                xl.Workbooks.Open(os.path.dirname(os.path.abspath(__file__))+"\\data.xlsx") 
                xs = xl.Worksheets("Data")
                xs.Columns.AutoFit()
                
            else:
                self.createPDF()
                
    def createPDF(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font('Arial', 'B', 18)
        pdf.cell(40, 10, 'Assets Inventory')
        pdf.ln(2)
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(40,30,'Company Name:')
        pdf.cell(80,30,'Audyne LLC')
        pdf.ln(2)
        pdf.cell(40,50,'Date:')
        d = datetime.datetime.now()
        pdf.cell(80,50,str(d.month)+"/"+str(d.day)+"/"+str(d.year))
        
        if self.tvwDatType.currentItem().parent().text(0)=="By Category":
            pdf.cell(100,50,'Category:' + self.tvwDatType.currentItem().text(0))
        else:
            pdf.cell(100,50,'Location:' + self.tvwDatType.currentItem().text(0))

        pdf.ln(30)
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(10, 10, 'No',1,0)
        pdf.cell(27, 10, '#Item Number',1,0)
        pdf.cell(27, 10, '#Serial Number',1,0)
        pdf.cell(30, 10, 'Description',1,0)
        pdf.cell(20, 10, 'Acq Date',1,0)
        pdf.cell(20, 10, 'Acq Cost',1,0)
        pdf.cell(20, 10, 'Depr Method',1,0)
        pdf.cell(20, 10, 'Usefull Live',1,0)
        pdf.cell(20, 10, 'Current Val',1,0)
        
        curLine=10
        pdf.ln(curLine)
        no=1
        
        pdf.set_font('Arial', '', 9)
        
        for x in range(self.tblAsset.rowCount()):
            
            pdf.cell(10, 10, str(no),1)
            pdf.cell(27, 10, self.tblAsset.item(x,0).text(),1)
            pdf.cell(27, 10, self.tblAsset.item(x,1).text(),1)
            pdf.cell(30, 10, self.tblAsset.item(x,2).text(),1)
            pdf.cell(20, 10, self.tblAsset.item(x,3).text(),1)
            pdf.cell(20, 10, self.tblAsset.item(x,4).text(),1,0,'R')
            pdf.cell(20, 10, self.tblAsset.item(x,5).text(),1)
            pdf.cell(20, 10, self.tblAsset.item(x,6).text(),1)
            pdf.cell(20, 10, self.tblAsset.item(x,7).text(),1,1,'R')
            no=no+1
            
        
        pdf.output('data.pdf')
        os.startfile(os.path.dirname(os.path.abspath(__file__))+"\\data.pdf")
        
    def fillGraph(self):
        cur = connection.connection()
        con = connection.con

        cur.execute("select tCategories.Name, count(tAssets.AssetID) from tCategories,tAssets where tAssets.CategoryID=tCategories.CategoryID group by tCategories.Name")
        rows = cur.fetchall()
        labelsv = []
        sizesv=[]
        for r in rows:
            labelsv.append(r[0])
            sizesv.append(r[1])
        self.ax1.clear()
        self.ax1.pie(sizesv,  labels=labelsv, autopct='%1.1f%%', shadow=True, startangle=140)
        self.ax1.axis('equal')
        self.canvCategory.draw_idle()
        
        cur.execute("select tLocations.LocName, count(tAssets.AssetID) from tLocations,tAssets where tAssets.LocationID=tLocations.LocationID group by tLocations.LocName")
        rows = cur.fetchall()
        labelsc = []
        sizesc=[]
        for r in rows:
            labelsc.append(r[0])
            sizesc.append(r[1])
        self.ax2.clear()
        self.ax2.pie(sizesc,  labels=labelsc, autopct='%1.1f%%', shadow=True, startangle=140)
        self.ax2.axis('equal')
        self.canvLocation.draw_idle()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = MainWin()
    sys.exit(app.exec_())
