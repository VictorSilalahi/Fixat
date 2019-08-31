from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtWidgets
from PyQt5.QtCore import QDate
import datetime
from inc import connection

class winAdd(QtWidgets.QDialog):
    def __init__(self,parent=None):
        super().__init__()
        self.parent = parent
        self.setWin()
    
    def setWin(self):
        self.setWindowTitle("Add new asset")
        self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)
    
        # layout seting
        self.winLayout = QtWidgets.QGridLayout(self)
    
        # set UI component
        self.lblAssetNo = QtWidgets.QLabel("Asset Number : ")
        self.lblAssetNo.setAlignment(QtCore.Qt.AlignRight)
        self.txtAssetNo = QtWidgets.QLineEdit()
        self.lblSN = QtWidgets.QLabel("Serial Number : ")
        self.lblSN.setAlignment(QtCore.Qt.AlignRight)
        self.txtSN = QtWidgets.QLineEdit()
        self.lblCategory = QtWidgets.QLabel("Asset Category : ")
        self.lblCategory.setAlignment(QtCore.Qt.AlignRight)
        self.cmbCategory = QtWidgets.QComboBox()
        self.lblLocation = QtWidgets.QLabel("Asset Location : ")
        self.lblLocation.setAlignment(QtCore.Qt.AlignRight)
        self.cmbLocation = QtWidgets.QComboBox()
        self.lblAssetDesc = QtWidgets.QLabel("Asset Description : ")
        self.lblAssetDesc.setAlignment(QtCore.Qt.AlignRight)
        self.txtAssetDesc = QtWidgets.QLineEdit()
        self.lblAcqDate = QtWidgets.QLabel("Acquaire Date : ")
        self.lblAcqDate.setAlignment(QtCore.Qt.AlignRight)
        self.dpAcqDate = QtWidgets.QDateEdit( QtCore.QDate.currentDate() )
        self.lblAcqCost = QtWidgets.QLabel("Acquaire Cost : ")
        self.lblAcqCost.setAlignment(QtCore.Qt.AlignRight)
        self.txtAcqCost = QtWidgets.QLineEdit()
        self.lblDep = QtWidgets.QLabel("Depreciation Method : ")
        self.lblDep.setAlignment(QtCore.Qt.AlignRight)
        self.cmbDep = QtWidgets.QComboBox()
        self.lblUseLive = QtWidgets.QLabel("Useful Live (years) : ")
        self.lblUseLive.setAlignment(QtCore.Qt.AlignRight)
        self.txtUseLive = QtWidgets.QLineEdit()
        self.lblWarning = QtWidgets.QLabel()
        self.lblWarning.setStyleSheet("color:rgb(200,0,0)")
        
        
        self.btnCancel= QtWidgets.QPushButton("Cancel")
        self.btnOK = QtWidgets.QPushButton("OK")
    
        # UI component position
        self.winLayout.addWidget(self.lblAssetNo,0,0)
        self.winLayout.addWidget(self.txtAssetNo,0,1)
        self.winLayout.addWidget(self.lblSN,0,2)
        self.winLayout.addWidget(self.txtSN,0,3)
        self.winLayout.addWidget(self.lblCategory,1,0)
        self.winLayout.addWidget(self.cmbCategory,1,1)
        self.winLayout.addWidget(self.lblLocation,1,2)
        self.winLayout.addWidget(self.cmbLocation,1,3)
        self.winLayout.addWidget(self.lblAssetDesc,2,0)
        self.winLayout.addWidget(self.txtAssetDesc,2,1,1,3)
        self.winLayout.addWidget(self.lblAcqDate,3,2)
        self.winLayout.addWidget(self.dpAcqDate,3,3)
        self.winLayout.addWidget(self.lblAcqCost,4,2)
        self.winLayout.addWidget(self.txtAcqCost,4,3)
        self.winLayout.addWidget(self.lblDep,5,2)
        self.winLayout.addWidget(self.cmbDep,5,3)
        self.winLayout.addWidget(self.lblUseLive,6,2)
        self.winLayout.addWidget(self.txtUseLive,6,3)
        self.winLayout.addWidget(self.lblWarning,7,0)
        

        self.winLayout.addWidget(self.btnCancel,8,2)
        self.winLayout.addWidget(self.btnOK,8,3)
        self.show()
    
        # event
        self.btnCancel.clicked.connect(self.closeWin)
        self.btnOK.clicked.connect(self.okAddEditAsset)
        self.fillComboList()
        
    def fillComboList(self):
        cur = connection.connection()
        row = cur.execute("select name from tCategories")
        for r in row:
            self.cmbCategory.addItem(r[0])
        row = cur.execute("select locname from tLocations")
        for r in row:
            self.cmbLocation.addItem(r[0])
        # depreciation method
        self.cmbDep.addItem("SLN")
        self.cmbDep.addItem("DDB")
        self.cmbDep.addItem("SYD")
        if self.parent.assetOp=="edit":
            self.editableData()
    
    def editableData(self):
        for i in range(self.parent.tblAsset.rowCount()):
            if self.parent.tblAsset.item(i,0).isSelected()== True:
                self.txtAssetNo.setText(self.parent.tblAsset.item(i,0).text())
                self.txtSN.setText(self.parent.tblAsset.item(i,1).text())
                
                cur = connection.connection()
                row = cur.execute("select tCategories.Name, tLocations.LocName, tAssets.AsDesc, tAssets.AcqCost, tAssets.DepMeth, tAssets.UsefulLive, tAssets.AcqDate  from tCategories,tLocations,tAssets where tCategories.CategoryID=tAssets.CategoryID and tLocations.LocationID=tAssets.LocationID and tAssets.AssetNo='"+self.parent.tblAsset.item(i,0).text()+"'")
                self.irow=i
                row = cur.fetchone()
                
                idx = self.cmbCategory.findText(row[0],  QtCore.Qt.MatchFixedString )
                if idx>=0:
                    self.cmbCategory.setCurrentIndex(idx)
                idx = self.cmbLocation.findText(row[1],  QtCore.Qt.MatchFixedString )
                if idx>=0:
                     self.cmbLocation.setCurrentIndex(idx)
                idx = self.cmbDep.findText(row[4],  QtCore.Qt.MatchFixedString )
                if idx>=0:
                     self.cmbDep.setCurrentIndex(idx)
                     
                self.txtAssetDesc.setText(row[2])
                self.txtAcqCost.setText(str(row[3]))
                self.txtUseLive.setText(str(row[5]))
                self.dpAcqDate.setDate(datetime.datetime.strptime(str(row[6]), '%m/%d/%Y'  ))
                
                break
        
        

    def closeWin(self):
        self.close()
    
    def okAddEditAsset(self):
        if (self.txtAssetNo.text()==""):
            self.lblWarning.setText("Please type asset number!")
            self.txtAssetNo.setFocus()
            return
        if (self.txtSN.text()==""):
            self.lblWarning.setText("Please type asset serial number!")
            self.txtSN.setFocus()
            return        if (self.txtAssetDesc.text()==""):
            self.lblWarning.setText("Please type asset Description!")
            self.txtAssetDesc.setFocus()
            return        if (self.txtAcqCost.text()==""):
            self.lblWarning.setText("Please type acquairing cost!")
            self.txtAcqCost.setFocus()
            return        if (self.txtAcqCost.text().isnumeric()==False):
            self.lblWarning.setText("Please type acquairing cost!")
            self.txtAcqCost.setFocus()
            return
        if (self.txtUseLive.text()==""):
            self.lblWarning.setText("Please type usefull year!")
            self.txtUseLive.setFocus()
            return
        if (self.txtUseLive.text().isnumeric()==False):
            self.lblWarning.setText("Please type usefull year!")
            self.txtUseLive.setFocus()
            return

        cur = connection.connection()
        con = connection.con

        if self.parent.assetOp=="add": 
            cur.execute("select count(*) as amt from tAssets where AssetNo='" + self.txtAssetNo.text() + "'")
            rsCount = cur.fetchone()
            if (rsCount[0]!=0):
                self.lblWarning.setText("Asset number exist in database. Please type new asset number!")
                self.txtAssetNo.setFocus()
                return
            cur.execute("select count(*) as amt from tAssets where SN='" + self.txtSN.text() + "'")
            rsCount = cur.fetchone()
            if (rsCount[0]!=0):
                self.lblWarning.setText("Serial number exist in database. Please type new serial number!")
                self.txtSN.setFocus()
                return
        
            cur.execute("select categoryID from tCategories where Name='" + self.cmbCategory.currentText() + "'")
            rsID = cur.fetchone()
            catID = rsID[0]
            
            cur.execute("select LocationID from tLocations where LocName='" + self.cmbLocation.currentText() + "'")
            rsID = cur.fetchone()
            locID = rsID[0]
    
            cur.execute( "insert into tAssets (CategoryID, LocationID, AssetNo, SN, AsDesc, AcqDate, AcqCost, DepMeth, UsefulLive) values (" + str(catID) + "," + str(locID) + ",'" + self.txtAssetNo.text() + "','" + self.txtSN.text() + "','" + self.txtAssetDesc.text() + "','" + self.dpAcqDate.text()  + "'," + self.txtAcqCost.text() + ",'" + self.cmbDep.currentText() + "'," + self.txtUseLive.text() + ")"  )
            con.commit()
            #if (self.cmbCategory.currentText()==self.parent.tvwDatType.currentItem().text(0)):
            self.parent.fillTable()
            self.close()
        else:
            cur.execute("select categoryID from tCategories where Name='" + self.cmbCategory.currentText() + "'")
            rsID = cur.fetchone()
            catID = rsID[0]
        
            cur.execute("select LocationID from tLocations where LocName='" + self.cmbLocation.currentText() + "'")
            rsID = cur.fetchone()
            locID = rsID[0]

            cur.execute( "update tAssets set CategoryID="+str(catID)+", LocationID="+str(locID)+", AssetNo='"+self.txtAssetNo.text()+"', SN='"+self.txtSN.text()+"', AsDesc='"+self.txtAssetDesc.text()+"', AcqDate='"+self.dpAcqDate.text()+"', AcqCost="+self.txtAcqCost.text()+", DepMeth='"+self.cmbDep.currentText()+"', UsefulLive="+self.txtUseLive.text()+" where AssetNo='"+self.parent.tblAsset.item(self.irow,0).text()+"'")
            con.commit()
            #if (self.cmbCategory.currentText()==self.parent.tvwDatType.currentItem().text(0)):
            self.parent.fillTable()
            self.close()
            
        
            