import sys

import pandas as pd

from mykiwoom.kiwoom import *
from config.myInfo import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic

form_class = uic.loadUiType("ui/main_window.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.kiwoom = Kiwoom()
        self.kiwoom.CommConnect()
        self.myInfo = MyInfo()

        # Timer1
        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        # Timer2
        self.timer2 = QTimer(self)
        self.timer2.start(1000 * 10)
        self.timer2.timeout.connect(self.timeout2)

        accounts_list = self.kiwoom.GetLoginInfo("ACCNO")
        self.comboBox_account.addItems(accounts_list)
        self.account = self.comboBox_account.currentText()
        self.secret = self.myInfo.account[self.account]

        self.lineEdit_code.textChanged.connect(self.code_changed)
        self.pushButton_order.clicked.connect(self.send_order)
        self.pushButton_check_balance.clicked.connect(self.check_balance)

        self.btn_file_load.clicked.connect(self.btn_file_load_clicked)

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = self.kiwoom.GetConnectState()
        if state == 1:
            state_msg = "서버 연결 중"
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def timeout2(self):
        if self.checkBox_check_balance_real.isChecked():
            self.check_balance()

    def code_changed(self):
        code = self.lineEdit_code.text()
        name = self.kiwoom.GetMasterCodeName(code)
        self.lineEdit_name.setText(name)

    def send_order(self):
        order_type_lookup = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4}
        hoga_lookup = {'지정가': "00", '시장가': "03"}

        account = self.comboBox_account.currentText()
        order_type = self.comboBox_order_type.currentText()
        code = self.lineEdit_code.text()
        hoga = self.comboBox_hoga.currentText()
        num = self.spinBox_num.value()

        if hoga == "시장가":
            price = 0
        else:
            price = self.spinBox_price.value()

        self.kiwoom.SendOrder("send_order_req", "0101", account, order_type_lookup[order_type], code, num, price,
                              hoga_lookup[hoga], "")

    def check_balance(self):
        # balance
        self.balance = pd.DataFrame()
        opw00001 = self.kiwoom.block_request("opw00001",
                                             계좌번호=self.account,
                                             비밀번호=self.secret,
                                             비밀번호입력매체구분="00",
                                             조회구분="1",
                                             output="예수금상세현황")

        opw00018 = self.kiwoom.block_request("opw00018",
                                             계좌번호=self.account,
                                             비밀번호=self.secret,
                                             비밀번호입력매체구분="00",
                                             조회구분="1",
                                             output="계좌평가결과")

        self.balance['d+2추정예수금'] = opw00001['d+2추정예수금'].apply(self.kiwoom.change_format)
        self.balance['총매입금액'] = opw00018['총매입금액'].apply(self.kiwoom.change_format)
        self.balance['총평가금액'] = opw00018['총평가금액'].apply(self.kiwoom.change_format)
        self.balance['총평가손익금액'] = opw00018['총평가손익금액'].apply(self.kiwoom.change_format)
        self.balance['총수익률(%)'] = opw00018['총수익률(%)'].apply(self.kiwoom.change_format2)
        self.balance['추정예탁자산'] = opw00018['추정예탁자산'].apply(self.kiwoom.change_format)

        for i in range(0, 6):
            item = QTableWidgetItem(self.balance.iloc[:,i][0])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_jango.setItem(0, i, item)

        # Item list
        self.item_list = pd.DataFrame()
        opw00018m = self.kiwoom.block_request("opw00018",
                                              계좌번호=self.account,
                                              비밀번호=self.secret,
                                              비밀번호입력매체구분="00",
                                              조회구분="1",
                                              output="계좌평가잔고개별합산")

        self.item_list['종목명'] = opw00018m['종목명'].apply(self.kiwoom.change_format)
        self.item_list['보유수량'] = opw00018m['보유수량'].apply(self.kiwoom.change_format)
        self.item_list['매입가'] = opw00018m['매입가'].apply(self.kiwoom.change_format)
        self.item_list['현재가'] = opw00018m['현재가'].apply(self.kiwoom.change_format)
        self.item_list['평가손익'] = opw00018m['평가손익'].apply(self.kiwoom.change_format)
        self.item_list['수익률(%)'] = opw00018m['수익률(%)'].apply(self.kiwoom.change_format2)

        self.item_list = self.item_list[self.item_list['종목명'] != '0']
        item_count = self.item_list.shape[0]
        self.tableWidget_stock.setRowCount(item_count)

        print(self.item_list)

        for j in range(item_count):
            row = self.item_list.iloc[j]
            for i in range(len(row)):
                item = QTableWidgetItem(row[i])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_stock.setItem(j, i, item)

        self.tableWidget_stock.resizeRowsToContents()


    def btn_file_load_clicked(self):
        fname = QFileDialog.getOpenFileName(self)
        # QMessageBox.about(self, "File loaded", fname[0])




if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    print(myWindow.account)
    print(myWindow.secret)
    # opw00001 = myWindow.kiwoom.block_request("opw00001",
    #                                          계좌번호=myWindow.account,
    #                                          비밀번호=myWindow.secret,
    #                                          비밀번호입력매체구분="00",
    #                                          조회구분="1",
    #                                          output="예수금상세현황")
    # myWindow.d2_deposit = myWindow.kiwoom.change_format(opw00001['d+2추정예수금'][0])
    # print(myWindow.d2_deposit)

    # opw00018 = myWindow.kiwoom.block_request("opw00018",
    #                                          계좌번호=myWindow.account,
    #                                          비밀번호=myWindow.secret,
    #                                          비밀번호입력매체구분="00",
    #                                          조회구분="1",
    #                                          output="계좌평가결과")

    # opw00018m = myWindow.kiwoom.block_request("opw00018",
    #                                           계좌번호=myWindow.account,
    #                                           비밀번호=myWindow.secret,
    #                                           비밀번호입력매체구분="00",
    #                                           조회구분="1",
    #                                           output="계좌평가잔고개별합산")
    # print(opw00018m)
    # print(opw00018m.shape)
    app.exec_()
