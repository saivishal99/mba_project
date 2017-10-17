import sys
import urllib2
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from xlrd import open_workbook

book = open_workbook("avg prices.xlsx")
sheet2 = book.sheet_by_index(1) # for sheet2
sheet6 = book.sheet_by_index(5) # for sheet6
columnA = []
columnB = []
columnC = []
columnD = []

for row in range(4,75):
	columnA.append(sheet2.cell(row, 2))
	columnB.append(sheet2.cell(row, 3))
	columnC.append(sheet2.cell(row, 4))
	columnD.append(sheet6.cell(row, 2))
#update = 0



string  = "Date: 04-Aug-2017,Tesla($80.00 per share),Tata($40.00 per share),Dabur($30.00 per share),CocaCola($42.00 per share)"
lst = string.split(',')
bght_lst = ['None']

app = QApplication(sys.argv)

class Form(QDialog):
	def __init__(self, parent=None):
		super(Form, self).__init__(parent)
		times = 0 
		self.date = lst[0]
		self.rates = lst[1:5]
		print self.rates
		print type(self.rates[0])
		print type(self.rates)
		print "Hello..."
		self.dateLabel = QLabel(self.date)
		self.fromComboBox = QComboBox()
	
		self.buyA_Button = QPushButton("Buy")
		self.buyB_Button = QPushButton("Buy")
		self.buyC_Button = QPushButton("Buy")
		self.buyD_Button = QPushButton("Buy")

		self.NumSharesA = QSpinBox()
		self.NumSharesB = QSpinBox()
		self.NumSharesC = QSpinBox()
		self.NumSharesD = QSpinBox()

		self.NumSharesA.setRange(1,1000)
		self.NumSharesB.setRange(1,1000)
		self.NumSharesC.setRange(1,1000)
		self.NumSharesD.setRange(1,1000)

		self.sellA_Button = QPushButton("Sell")
		self.sellB_Button = QPushButton("Sell")
		self.sellC_Button = QPushButton("Sell")
		self.sellD_Button = QPushButton("Sell")

		self.share1=QLabel("A")
#		self.im1 =  QImage("PL_gragh.png")
		self.share2=QLabel("B")
		self.share3=QLabel("C")
		self.share4=QLabel("D")

		self.Share1=QLabel("A:	")
                self.Share2=QLabel("B:	")
                self.Share3=QLabel("C:	")
		self.Share4=QLabel("D:  ")
		
		self.state1=QLabel("Rising")
		self.state2=QLabel("Falling")
		self.state3=QLabel("Rising")
		self.state4=QLabel("Falling")

		self.price1=QLabel("$20.00")
		self.price2= QLabel("$20.00")
		self.price3= QLabel("$20.00")
		self.price4= QLabel("$20.00")
		
		self.HistLabel = QLabel("History")
		self.HistoryComboBox = QComboBox()
		self.Cash = QLabel("$10000.00")
		self.CashLbl = QLabel("Cash: ")
	
		self.PortfolioVal = QLabel("$0.00")
		self.PortfolioValLbl = QLabel("PV: ")

		self.holdingLbl = QLabel("Holding: ")
		self.holdA = QLabel("A")
		self.holdB = QLabel("B")
		self.holdC = QLabel("C")
		self.holdD = QLabel("D")
		self.holdNumA = QLabel("0 shares")
                self.holdNumB = QLabel("0 shares")
                self.holdNumC = QLabel("0 shares")
		self.holdNumD = QLabel("0 shares")
		self.setWindowTitle("StockPlay")

		self.Clock = QLabel("00:00")
		self.ClockLbl = QLabel("CLOCK: ")
		
		Grid = QGridLayout()
		Grid.addWidget(self.share1,1,1)
		Grid.addWidget(self.share2,4,1)
		Grid.addWidget(self.share3,7,1)
		Grid.addWidget(self.share4,10,1)
		Grid.addWidget(self.NumSharesA,1,6)
		Grid.addWidget(self.NumSharesB,3,6)
		Grid.addWidget(self.NumSharesC,5,6)
		Grid.addWidget(self.NumSharesD,7,6)
		Grid.addWidget(self.buyA_Button,1,7)
		Grid.addWidget(self.buyB_Button,3,7)
		Grid.addWidget(self.buyC_Button,5,7)
		Grid.addWidget(self.buyD_Button,7,7)
#		Grid.addWidget(self.buyD_Button,1,7)
		Grid.addWidget(self.sellA_Button,1,8)
		Grid.addWidget(self.sellB_Button,3,8)
		Grid.addWidget(self.sellC_Button,5,8)
		Grid.addWidget(self.sellD_Button,7,8)
		Grid.addWidget(self.price1,2,0)
		Grid.addWidget(self.price2,5,0)
		Grid.addWidget(self.price3,8,0)
		Grid.addWidget(self.price4,11,0)
		Grid.addWidget(self.state1,2,2)
		Grid.addWidget(self.state2,5,2)
		Grid.addWidget(self.state3,8,2)
		Grid.addWidget(self.state4,11,2)
		Grid.addWidget(self.ClockLbl,0,4)
		Grid.addWidget(self.Clock,0,5)
		Grid.addWidget(self.Share1,1,5)
		Grid.addWidget(self.Share2,3,5)
		Grid.addWidget(self.Share3,5,5)
		Grid.addWidget(self.Share4,7,5)
		Grid.addWidget(self.HistLabel,7,5,7,6)
		Grid.addWidget(self.HistoryComboBox,7,6,7,7)
		Grid.addWidget(self.holdingLbl,12,6)
		Grid.addWidget(self.holdA,12,7)
		Grid.addWidget(self.holdB,12,8)
		Grid.addWidget(self.holdC,12,9)
		Grid.addWidget(self.holdD,12,10)
		Grid.addWidget(self.holdNumA,13,7)
		Grid.addWidget(self.holdNumB,13,8)
		Grid.addWidget(self.holdNumC,13,9)
		Grid.addWidget(self.holdNumD,13,10)
		Grid.addWidget(self.CashLbl,17,8)
		Grid.addWidget(self.Cash,17,9)
		Grid.addWidget(self.PortfolioValLbl,17,6)
		Grid.addWidget(self.PortfolioVal,17,7)
		self.setLayout(Grid)		
		
		self.connect(self.buyA_Button, SIGNAL("clicked()"), self.updateBuyAUi)
		self.connect(self.buyB_Button, SIGNAL("clicked()"), self.updateBuyBUi)
		self.connect(self.buyC_Button, SIGNAL("clicked()"), self.updateBuyCUi)
		self.connect(self.buyD_Button, SIGNAL("clicked()"), self.updateBuyDUi)
		self.connect(self.sellA_Button, SIGNAL("clicked()"), self.updateSellAUi)
		self.connect(self.sellB_Button, SIGNAL("clicked()"), self.updateSellBUi)
		self.connect(self.sellC_Button, SIGNAL("clicked()"), self.updateSellCUi)
		self.connect(self.sellD_Button, SIGNAL("clicked()"), self.updateSellDUi)




	###############################################################################################
	
	def updateBuyAUi(self):
		doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc = QTextDocument()
                doc.setHtml(self.price1.text())
                text2 = doc.toPlainText()
		curr, tcost = text2.split("$")
		cost = float(tcost)
		numShrs = self.NumSharesA.value()
		doc.setHtml(self.Cash.text())
		text = doc.toPlainText()
		curr, tcash = text.split("$")
		cash = float(tcash)
		if cash < cost*numShrs:
			label = QLabel("<font color=red size=5><b>" + "You don't have enough money!" + "</b>" + 
					"(Click on this and then click on GameBox)"+"</font>")
			label.setWindowFlags(Qt.SplashScreen)
			label.show()
			app1.QApplication(sys.argv)
			app1._exec()
			app1.close()
			return
		cash = cash - numShrs*cost
		# Updating the HoldingLbl for A
		doc3 = QTextDocument()
                doc3.setHtml(self.holdNumA.text())
                text3 = doc3.toPlainText()
		tval1, tval2 = text3.split(" ")
		val1 = int(tval1) + numShrs
		self.holdNumA.setText(str(val1)+" shares")
		self.Cash.setText("$"+str(cash))
		# INSERT INTO HISTORY
		self.HistoryComboBox.insertItem(0,"Spent:"+str(cost*numShrs)+" For:ShareA("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
	#####################################################################################################



	#####################################################################################################
	def updateBuyBUi(self):
                doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc = QTextDocument()
                doc.setHtml(self.price2.text())
                text2 = doc.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                numShrs = self.NumSharesB.value()
                doc.setHtml(self.Cash.text())
                text = doc.toPlainText()
                curr, tcash = text.split("$")
                cash = float(tcash)
                if cash < cost*numShrs:
                        label = QLabel("<font color=red size=5><b>" + "You don't have enough money!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return
                cash = cash - numShrs*cost
                # Updating the HoldingLbl for A
                doc3 = QTextDocument()
                doc3.setHtml(self.holdNumB.text())
                text3 = doc3.toPlainText()
                tval1, tval2 = text3.split(" ")
                val1 = int(tval1) + numShrs
                self.holdNumB.setText(str(val1)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Spent:"+str(cost*numShrs)+" For:ShareB("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
        #####################################################################################################
	

	#####################################################################################################
	def updateBuyCUi(self):
                doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc = QTextDocument()
                doc.setHtml(self.price3.text())
                text2 = doc.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                numShrs = self.NumSharesC.value()
                doc.setHtml(self.Cash.text())
                text = doc.toPlainText()
                curr, tcash = text.split("$")
                cash = float(tcash)
                if cash < cost*numShrs:
                        label = QLabel("<font color=red size=5><b>" + "You don't have enough money!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return
                cash = cash - numShrs*cost
                # Updating the HoldingLbl for C
                doc3 = QTextDocument()
                doc3.setHtml(self.holdNumC.text())
                text3 = doc3.toPlainText()
                tval1, tval2 = text3.split(" ")
                val1 = int(tval1) + numShrs
                self.holdNumC.setText(str(val1)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Spent:"+str(cost*numShrs)+" For:ShareC("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
        #####################################################################################################


	######################################################################################################
	def updateBuyDUi(self):
                doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc = QTextDocument()
                doc.setHtml(self.price4.text())
                text2 = doc.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                numShrs = self.NumSharesD.value()
                doc.setHtml(self.Cash.text())
                text = doc.toPlainText()
                curr, tcash = text.split("$")
                cash = float(tcash)
                if cash < cost*numShrs:
                        label = QLabel("<font color=red size=5><b>" + "You don't have enough money!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return
                cash = cash - numShrs*cost
                # Updating the HoldingLbl for D
                doc3 = QTextDocument()
                doc3.setHtml(self.holdNumD.text())
                text3 = doc3.toPlainText()
                tval1, tval2 = text3.split(" ")
                val1 = int(tval1) + numShrs
                self.holdNumD.setText(str(val1)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Spent:"+str(cost*numShrs)+" For:ShareD("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
        #####################################################################################################



	#####################################################################################################
	def updateSellAUi(self):
		doc = QTextDocument()
                doc.setHtml(self.holdNumA.text())
                text = doc.toPlainText()
		num1, num2 = text.split(" ")
		numShrs = self.NumSharesA.value()
		if numShrs > float(num1):
			label = QLabel("<font color=red size=5><b>" + "You don't have enough shares!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return

		doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc2 = QTextDocument()
                doc2.setHtml(self.price1.text())
                text2 = doc2.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                doc4 = QTextDocument()
		doc4.setHtml(self.Cash.text())
                text4 = doc4.toPlainText()
                curr, tcash = text4.split("$")
                cash = float(tcash)
                cash = cash + numShrs*cost
		 # Updating the HoldingLbl for A
                self.holdNumA.setText(str(int(num1)-numShrs)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Earned:"+str(cost*numShrs)+" For:ShareA("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
	##############################################################################################################


	########################################################################################################
	def updateSellBUi(self):
                doc = QTextDocument()
                doc.setHtml(self.holdNumB.text())
                text = doc.toPlainText()
                num1, num2 = text.split(" ")
                numShrs = self.NumSharesB.value()
                if numShrs > float(num1):
                        label = QLabel("<font color=red size=5><b>" + "You don't have enough shares!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return

                doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc2 = QTextDocument()
                doc2.setHtml(self.price2.text())
                text2 = doc2.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                doc4 = QTextDocument()
                doc4.setHtml(self.Cash.text())
                text4 = doc4.toPlainText()
                curr, tcash = text4.split("$")
                cash = float(tcash)
                cash = cash + numShrs*cost
                 # Updating the HoldingLbl for A
                self.holdNumB.setText(str(int(num1)-numShrs)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Earned:"+str(cost*numShrs)+" For:ShareB("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
        ##############################################################################################################


	#############################################################################################################
	def updateSellCUi(self):
                doc = QTextDocument()
                doc.setHtml(self.holdNumC.text())
                text = doc.toPlainText()
                num1, num2 = text.split(" ")
                numShrs = self.NumSharesC.value()
                if numShrs > float(num1):
                        label = QLabel("<font color=red size=5><b>" + "You don't have enough shares!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return

                doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc2 = QTextDocument()
                doc2.setHtml(self.price3.text())
                text2 = doc2.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                doc4 = QTextDocument()
                doc4.setHtml(self.Cash.text())
                text4 = doc4.toPlainText()
                curr, tcash = text4.split("$")
                cash = float(tcash)
                cash = cash + numShrs*cost
                 # Updating the HoldingLbl for A
                self.holdNumC.setText(str(int(num1)-numShrs)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Earned:"+str(cost*numShrs)+" For:ShareC("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
        ##############################################################################################################
	

	##############################################################################################################
	def updateSellDUi(self):
                doc = QTextDocument()
                doc.setHtml(self.holdNumD.text())
                text = doc.toPlainText()
                num1, num2 = text.split(" ")
                numShrs = self.NumSharesD.value()
                if numShrs > float(num1):
                        label = QLabel("<font color=red size=5><b>" + "You don't have enough shares!" + "</b>" +
                                        "(Click on this and then click on GameBox)"+"</font>")
                        label.setWindowFlags(Qt.SplashScreen)
                        label.show()
                        app1.QApplication(sys.argv)
                        app1._exec()
                        app1.close()
                        return

                doc1 = QTextDocument()
                doc1.setHtml(self.Clock.text())
                text1 = doc1.toPlainText()
                doc2 = QTextDocument()
                doc2.setHtml(self.price4.text())
                text2 = doc2.toPlainText()
                curr, tcost = text2.split("$")
                cost = float(tcost)
                doc4 = QTextDocument()
                doc4.setHtml(self.Cash.text())
                text4 = doc4.toPlainText()
                curr, tcash = text4.split("$")
                cash = float(tcash)
                cash = cash + numShrs*cost
                 # Updating the HoldingLbl for A
                self.holdNumD.setText(str(int(num1)-numShrs)+" shares")
                self.Cash.setText("$"+str(cash))
                # INSERT INTO HISTORY
                self.HistoryComboBox.insertItem(0,"Earned:"+str(cost*numShrs)+" For:ShareD("+text2+")"+" At-"+text1+" Holding:"+"$"+str(cash))
        ##############################################################################################################



	def updateTime(self):
		doc = QTextDocument()
		doc.setHtml(self.Clock.text())
	  	text = doc.toPlainText()
		tmin, tsec = text.split(":")
		secs = int(tsec)
		mins = int(tmin)
		secs += 1
		if(secs == 60):
			tsec = "00"
			tmin = str(mins+1)
		else:
			tsec = str(secs)
			tmin = str(mins)
		self.Clock.setText(tmin+":"+tsec)
#		if update % 2 == 0:
#			pr1 = next(columnA)
#			pr2 = next(columnB)
#			pr3 = next(columnC)
#		  	pr4 = next(columnD)
#		  	self.price1.setText("$"+str(pr1))
#		 	self.price2.setText("$"+str(pr2))
#			self.price3.setText("$"+str(pr3))
#			self.price4.setText("$"+str(pr4))
#		++update
		st, pr1 = str(columnA[int(secs/2)]).split(":")
		st, pr2 = str(columnB[int(secs/2)]).split(":")
		st, pr3 = str(columnC[int(secs/2)]).split(":")
		st, pr4 = str(columnD[int(secs/2)]).split(":")

		self.price1.setText("$"+str(pr1))
		self.price2.setText("$"+str(pr2))
		self.price3.setText("$"+str(pr3))
                self.price4.setText("$"+str(pr4))








form = Form()
form.show()
t = 0
for i in range(120): # game span of 120 secs
	QTimer.singleShot(t+1000, form.updateTime) # 2 seconds
	t += 1000
app.exec_()



