import sys


from mykiwoom.kiwoom import *
from config.myInfo import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from PyQt5.QtTest import *  # QTest.qWait(5000)
import numpy as np

import requests
import pandas as pd
import python_quant_super_value as pq
import time
import bs4
import urllib.parse
import matplotlib.pyplot as plt
import numpy
from pykrx import stock
from datetime import datetime

form_class = uic.loadUiType("ui/main_window.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.kiwoom = Kiwoom()
        self.kiwoom.CommConnect()
        self.myInfo = MyInfo()

        if self.kiwoom.GetLoginInfo("GetServerGubun") == "1":
            self.server_gubun = "모의투자"
        else:
            self.server_gubun = "실제운영"

        # portfolio file
        self.account = None
        self.secret = None
        self.fname = None
        self.fname_pf = None
        self.total_list = None
        self.portfolio_check = None
        self.my_comment = "Ready"

        # Timer1
        self.timer = QTimer(self)
        self.timer.start(100)
        self.timer.timeout.connect(self.timeout)

        # Timer2
        self.timer2 = QTimer(self)
        self.timer2.start(1000 * 10)
        self.timer2.timeout.connect(self.timeout2)

        # Data
        self.balance = pd.DataFrame()
        self.item_list = pd.DataFrame()
        self.portfolio = pd.DataFrame()

        accounts_list = self.kiwoom.GetLoginInfo("ACCNO")
        self.comboBox_account.addItems(accounts_list)
        self.account_update()

        QTest.qWait(1000)

        self.lineEdit_code.textChanged.connect(self.code_changed)
        self.pushButton_order.clicked.connect(self.send_order)
        self.pushButton_check_balance.clicked.connect(self.check_balance)

        self.btn_file_load.clicked.connect(self.btn_file_load_clicked)
        self.pushButton_order_auto.clicked.connect(self.send_order_auto)
        self.pushButton_get_portfolio.clicked.connect(self.get_portfolio)

    def account_update(self):
        self.account = self.comboBox_account.currentText()
        try:
            self.secret = self.myInfo.account[self.account]
        except KeyError:
            self.secret = '0000'
        print("계좌번호 : {}".format(self.account))

    def timeout(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = self.kiwoom.GetConnectState()
        if state == 1:
            state_msg = "서버 연결 중"
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg + " | " + self.my_comment)

    def timeout2(self):
        if self.checkBox_check_balance_real.isChecked():
            self.check_balance()

    def code_changed(self):
        code = self.lineEdit_code.text()
        name = self.kiwoom.GetMasterCodeName(code)
        self.lineEdit_name.setText(name)

    def get_code6(self, code):
        if len(code) == 7:
            code = code[1:]
        return code

    def get_master_code_name(self, code):
        if len(code) == 7:
            code = code[1:]
        return self.kiwoom.GetMasterCodeName(code)

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
        print("{} {}주 {}주문 완료".format(self.lineEdit_name.text(), num, self.comboBox_order_type.currentText()))
        self.my_comment = "{} {}주 {}주문 완료".format(self.lineEdit_name.text(), num, self.comboBox_order_type.currentText())

    def fill_QTable(self, dat, qtable, vertical=False):
        df = dat.copy()
        for i in range(df.shape[1]):
            if df.iloc[:, i].name == '종목코드':
                continue
            elif df.iloc[:, i].name.find('률') > 0:
                df.iloc[:, i] = df.iloc[:, i].apply(self.change_format2)
            else:
                df.iloc[:, i] = df.iloc[:, i].apply(self.change_format)

        if vertical:
            row_count = df.shape[0]
            qtable.setColumnCount(row_count)
            for j in range(row_count):
                row = df.iloc[j]
                for i in range(len(row)):
                    item = QTableWidgetItem(row[i])
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                    qtable.setItem(i, j, item)
        else:
            row_count = df.shape[0]
            qtable.setRowCount(row_count)
            for j in range(row_count):
                row = df.iloc[j]
                for i in range(len(row)):
                    item = QTableWidgetItem(row[i])
                    item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                    qtable.setItem(j, i, item)
            qtable.resizeRowsToContents()

    def check_balance(self):
        self.account_update()

        # balance
        self.my_comment = "잔고 조회 중..."
        opw00001 = self.my_block_request("opw00001",
                                         계좌번호=self.account,
                                         비밀번호=self.secret,
                                         비밀번호입력매체구분="00",
                                         조회구분="1",
                                         output="예수금상세현황")
        opw00018 = self.my_block_request("opw00018",
                                         계좌번호=self.account,
                                         비밀번호=self.secret,
                                         비밀번호입력매체구분="00",
                                         조회구분="1",
                                         output="계좌평가결과")
        balance_merge = pd.concat([opw00001[['d+2추정예수금']],
                                   opw00018.iloc[[0]][['총매입금액', '총평가금액', '총평가손익금액', '총수익률(%)', '추정예탁자산']]],
                                  axis=1)
        if self.server_gubun == "실제운영":
            balance_merge['총수익률(%)'] = balance_merge['총수익률(%)'] / 100
        self.balance = balance_merge
        self.fill_QTable(self.balance, self.tableWidget_jango, vertical=True)
        self.my_comment = "잔고 조회 완료, 보유종목현황 조회 중..."

        # Item list
        opw00018m = self.my_block_request("opw00018",
                                          계좌번호=self.account,
                                          비밀번호=self.secret,
                                          비밀번호입력매체구분="00",
                                          조회구분="1",
                                          output="계좌평가잔고개별합산")

        opw00018m['종목코드'] = opw00018m['종목번호'].apply(self.get_code6)
        self.item_list = opw00018m[['종목코드', '종목명', '보유수량', '매입가', '현재가', '평가손익', '수익률(%)']]
        if self.server_gubun == "실제운영":
            self.item_list['수익률(%)'] = self.item_list['수익률(%)'] / 100
        self.fill_QTable(self.item_list, self.tableWidget_stock)
        self.my_comment = "잔고 및 보유종목현황 조회 완료"

        return self.balance, self.item_list

    def check_portfolio(self):
        self.my_comment = "포트폴리오 불러오는 중..."
        self.portfolio = pd.read_excel(self.fname[0])
        self.portfolio['종목코드'] = self.portfolio['종목코드'].apply(self.get_code6)
        self.portfolio['기업명'] = self.portfolio['종목코드'].apply(self.get_master_code_name)
        self.portfolio = self.portfolio[['종목코드', '기업명']]
        self.portfolio.columns = ['종목코드', '종목명']
        self.portfolio = pd.merge(self.portfolio, self.item_list[['종목명', '보유수량']],
                                  how='left', on='종목명')
        self.portfolio['보유수량'] = self.portfolio['보유수량'].replace(np.nan, 0).astype('int')

        optkwfid = self.my_block_request("OPTKWFID",
                                         종목코드=';'.join(self.portfolio['종목코드']),
                                         output="관심종목정보")
        self.portfolio = pd.merge(self.portfolio, optkwfid[['종목코드', '현재가']],
                                  how='left', on='종목코드')

        self.portfolio['현재가'] = abs(self.portfolio['현재가'].astype('int'))
        self.portfolio['평가금액'] = self.portfolio['보유수량'] * self.portfolio['현재가']
        if self.portfolio_check is None:
            self.portfolio_check = self.portfolio[['종목코드']].copy()
            self.portfolio_check['완료'] = ""
        self.portfolio['구분'] = ""
        self.portfolio['매매수량'] = ""
        self.portfolio = pd.merge(self.portfolio, self.portfolio_check,
                                  how='left', on='종목코드')
        max_row = int(self.lineEdit_rank_meme.text())
        self.portfolio = self.portfolio.iloc[:max_row]
        print(self.portfolio)
        self.fill_QTable(self.portfolio, self.tableWidget_portfolio)
        self.my_comment = "포트폴리오 조회 완료"
        return self.portfolio

    def btn_file_load_clicked(self):
        self.fname = QFileDialog.getOpenFileName(self)
        self.check_balance()
        self.check_portfolio()

    def send_order_auto(self):
        order_type_auto = self.comboBox_order_type_auto.currentText()
        if self.fname == None:
            self.fname = QFileDialog.getOpenFileName(self)
        self.check_balance()
        self.check_portfolio()
        self.my_comment = "자동매매 시작"
        price_per_stock = self.balance['추정예탁자산'][0] / self.portfolio.shape[0]

        cols = ['종목코드', '종목명', '보유수량']
        self.total_list = pd.concat([self.item_list[cols], self.portfolio[cols]]).drop_duplicates()

        optkwfid = self.my_block_request("OPTKWFID",
                                         종목코드=';'.join(self.total_list['종목코드']),
                                         output="관심종목정보")

        self.total_list = pd.merge(self.total_list, optkwfid[['종목코드', '현재가']],
                                   how='left', on='종목코드')

        self.total_list['현재가'] = abs(self.total_list['현재가'].astype('int'))
        self.total_list['평가금액'] = self.total_list['보유수량'] * self.total_list['현재가']

        self.total_list['구분'] = np.where(~self.total_list['종목코드'].isin(self.portfolio['종목코드']), '전량매도',
                                         np.where(self.total_list['평가금액'] > price_per_stock,
                                                  '부분매도', '매수'))

        self.total_list['매매금액'] = np.where(self.total_list['구분'] == '전량매도', self.total_list['평가금액'],
                                           abs(round(price_per_stock - self.total_list['평가금액'], 0).astype('int')))
        self.total_list['매매수량'] = np.where(self.total_list['구분'] == '전량매도', self.total_list['보유수량'],
                                           round(self.total_list['매매금액'] / self.total_list['현재가'], 0).astype('int'))
        self.total_list['최대매매수량'] = round(self.spinBox_price_auto.value() * 10000 / self.total_list['현재가'], 0).astype('int')
        self.total_list['최종매매수량'] = np.where(self.total_list['매매수량'] <= self.total_list['최대매매수량'],
                                             self.total_list['매매수량'], self.total_list['최대매매수량'])
        self.total_list = pd.merge(self.total_list, self.portfolio_check, how='left', on='종목코드')
        self.total_list['완료'] = self.total_list['완료'].replace(np.nan, "")
        self.total_list.to_excel("temp.xlsx")

        if order_type_auto in ['매도', '매도매수']:
            # 1. 보유종목 중 포트폴리오 미포함 종목 매도
            self.my_comment = "자동매매 : 보유종목 중 포트폴리오 미포함 종목 매도 중..."
            QTest.qWait(1000)
            sell_list = self.total_list.loc[(self.total_list['구분'] == '전량매도') & (self.total_list['완료'] != '완료')]

            if len(sell_list) != 0:
                for i in range(0, sell_list.shape[0]):
                    code = sell_list['종목코드'].iloc[i]
                    num = int(sell_list['최종매매수량'].iloc[i])
                    if num > 0:
                        self.kiwoom.SendOrder("send_order_req", "0101", self.account, 2, code, num, 0, "03", "")
                        print("{} {}주 매도주문 완료".format(sell_list['종목명'].iloc[i], num))
                    if sell_list['매매수량'].iloc[i] == sell_list['최종매매수량'].iloc[i]:
                        self.total_list.loc[self.total_list['종목코드'] == code, '완료'] = '완료'

            # 2. 포트폴리오 종목 중 기준가 초과 종목 매도
            self.my_comment = "자동매매 : 포트폴리오 종목 중 기준가 초과 종목 매도 중..."
            QTest.qWait(1000)
            sell_list2 = self.total_list.loc[(self.total_list['구분'] == '부분매도') & (self.total_list['완료'] != '완료')]

            if len(sell_list2) != 0:
                for i in range(0, sell_list2.shape[0]):
                    code = sell_list2['종목코드'].iloc[i]
                    num = int(sell_list2['최종매매수량'].iloc[i])
                    if num > 0:
                        self.kiwoom.SendOrder("send_order_req", "0101", self.account, 2, code, num, 0, "03", "")
                        print("{} {}주 매도주문 완료".format(sell_list2['종목명'].iloc[i], num))
                    if sell_list2['매매수량'].iloc[i] == sell_list2['최종매매수량'].iloc[i]:
                        self.portfolio_check.loc[self.portfolio_check['종목코드'] == code, '완료'] = '완료'
                        self.portfolio.loc[self.portfolio['종목코드'] == code, '완료'] = '완료'
                        self.total_list.loc[self.total_list['종목코드'] == code, '완료'] = '완료'

        if order_type_auto in ['매도매수']:
            self.my_comment = "자동매매 : 포트폴리오 종목 중 기준가 미달 종목 매수 전 10초 대기 중..."
            QTest.qWait(10000)

        if order_type_auto in ['매수', '매도매수']:
            # 3. 포트폴리오 종목 중 기준가 미달 종목 매수
            self.my_comment = "자동매매 : 포트폴리오 종목 중 기준가 미달 종목 매수 중..."
            QTest.qWait(1000)
            buy_list = self.total_list.loc[(self.total_list['구분'] == '매수') & (self.total_list['완료'] != '완료')]
            if len(buy_list) != 0:
                for i in range(0, buy_list.shape[0]):
                    code = buy_list['종목코드'].iloc[i]
                    num = int(buy_list['최종매매수량'].iloc[i])
                    if num > 0:
                        self.kiwoom.SendOrder("send_order_req", "0101", self.account, 1, code, num, 0, "03", "")
                        print("{} {}주 매수주문 완료".format(buy_list['종목명'].iloc[i], num))
                    if buy_list['매매수량'].iloc[i] == buy_list['최종매매수량'].iloc[i]:
                        self.portfolio_check.loc[self.portfolio_check['종목코드'] == code, '완료'] = '완료'
                        self.portfolio.loc[self.portfolio['종목코드'] == code, '완료'] = '완료'
                        self.total_list.loc[self.total_list['종목코드'] == code, '완료'] = '완료'

        self.fill_QTable(self.total_list[['종목코드', '종목명', '보유수량', '현재가', '평가금액', '구분', '최종매매수량', '완료']],
                         self.tableWidget_portfolio)

        self.portfolio_check = self.total_list[['종목코드', '완료']].copy()
        self.my_comment = "자동매매 완료: 전량매도({} / {}), 부분매도({} / {}), 매수({} / {})".format(
            self.total_list.loc[(self.total_list['구분'] == '전량매도') & (self.total_list['완료'] == '완료')].shape[0],
            self.total_list.loc[(self.total_list['구분'] == '전량매도')].shape[0],
            self.total_list.loc[(self.total_list['구분'] == '부분매도') & (self.total_list['완료'] == '완료')].shape[0],
            self.total_list.loc[(self.total_list['구분'] == '부분매도')].shape[0],
            self.total_list.loc[(self.total_list['구분'] == '매수') & (self.total_list['완료'] == '완료')].shape[0],
            self.total_list.loc[(self.total_list['구분'] == '매수')].shape[0])

    def get_portfolio(self):
        self.fname_pf = QFileDialog.getOpenFileName(self)

        cap_high = float(self.lineEdit_high.text())
        cap_low = float(self.lineEdit_low.text())
        portfolio_rank = int(self.lineEdit_rank.text())

        # 전종목 시가총액 구하기 (https://yobro.tistory.com/142)
        today = datetime.today().strftime("%Y%m%d")
        df_kospi = stock.get_market_cap_by_ticker(today, market='KOSPI')
        df_kosdaq = stock.get_market_cap_by_ticker(today, market='KOSDAQ')
        code_data = pd.concat([df_kospi, df_kosdaq], axis=0)
        code_data['기업명'] = [stock.get_market_ticker_name(ticker) for ticker in code_data.index]
        code_data = code_data[code_data.거래량 > 0]  # 거재정지 종목 제외
        code_data = code_data[['기업명', '종가', '시가총액']]
        code_data.index = 'A' + code_data.index
        code_data.index.name = '종목코드'

        # 시총 하위 20%
        def get_cap_range(value_df, ratio=[0.8, 1.0]):
            self.my_comment = "포트폴리오 분석 시작"
            QTest.qWait(1000)
            temp_df = value_df
            temp_df['시가총액'] = pd.to_numeric(temp_df['시가총액'])
            temp_df = temp_df.sort_values(by='시가총액', ascending=False)
            temp_df['시가총액비율'] = temp_df['시가총액'].rank(ascending=False) / temp_df.shape[0]
            sorted_cap_value = temp_df[(temp_df['시가총액비율'] > min(ratio)) & (temp_df['시가총액비율'] <= max(ratio))]
            return sorted_cap_value

        code_data_filter = get_cap_range(code_data, ratio=[cap_high, cap_low])
        code_data_filter = code_data_filter[['기업명']]

        # fnguide 데이터 수집하기
        total_value = pd.DataFrame()
        for num, code in enumerate(code_data_filter.index):
            if num % 1 == 0:
                print("포트폴리오 분석 중 : [{} / {}] {} {}".format(
                    num + 1, code_data_filter.shape[0], code[1:], code_data.iloc[num, 0]))
                self.my_comment = "포트폴리오 분석 중 : [{} / {}] {} {}".format(
                    num + 1, code_data_filter.shape[0], code[1:], code_data.iloc[num, 0])

            try:
                QTest.qWait(1000)
                try:
                    cap = pq.make_cap_dataframe(code)
                    fs_df = pq.make_fs_dataframe(code)
                    fhd_df = pq.make_fhd_dataframe(code)
                except requests.exceptions.Timeout:
                    time.sleep(60)
                    cap = pq.make_cap_dataframe(code)
                    fs_df = pq.make_fs_dataframe(code)
                    fhd_df = pq.make_fhd_dataframe(code)
                except ValueError:
                    continue
                except KeyError:
                    continue
                except UnboundLocalError:
                    continue

                value = pd.merge(cap, fs_df, how='outer', right_index=True, left_index=True)
                value = pd.merge(value, fhd_df, how='outer', right_index=True, left_index=True)
                value['1/PSR'] = float(value['매출액'][0]) / float(value['시가총액'][0])
                value['1/PER'] = float(value['순이익'][0]) / float(value['시가총액'][0])
                value['1/PBR'] = float(value['자본'][0]) / float(value['시가총액'][0])
                value['1/PCR'] = float(value['영업현금흐름'][0]) / float(value['시가총액'][0])
                value = pd.merge(code_data_filter.loc[[code]], value, how='outer', right_index=True, left_index=True)
                if num == 0:
                    total_value = value
                else:
                    total_value = pd.concat([total_value, value])

                if num + 1 == code_data_filter.shape[0]:
                    print("Analysis Done")

            except ValueError:
                continue
            except KeyError:
                continue
            except UnboundLocalError:
                continue
        total_value['적정가'] = round((total_value['BPS'].astype('float') * (1 + total_value['ROE3년'] * 0.01 - 0.05 * (
                (100 + total_value['부채비율3년']) / 100) ** 0.5) ** 10).astype(float), -1)

        # # super value 전략 구현하기
        value_list = ['1/PSR', '1/PER', '1/PBR', '1/PCR']
        value_df = total_value
        value_combo = pq.value_combo(value_df, value_list, portfolio_rank)
        value_combo.to_excel(self.fname_pf[0])
        self.my_comment = "포트폴리오 분석 완료 (경로 : {})".format(self.fname_pf[0])

    def convert_type(self, x):
        if x.name != '종목코드':
            try:
                x = x.astype('int')
            except ValueError:
                x = x.astype('float', errors='ignore')
        return x

    def change_format(self, data):
        data = str(data)
        if data == '':
            strip_data = ''
        else:
            strip_data = data.lstrip('-0')
            if strip_data == '':
                strip_data = '0'

        try:
            format_data = format(int(strip_data), ',d')
        except:
            try:
                format_data = format(float(strip_data))
            except:
                format_data = format(strip_data)

        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    def change_format2(self, data):  # 수익률 포맷
        data = str(data)
        if data == '':
            strip_data = ''
        else:
            strip_data = data.lstrip('-0')
            if strip_data == '':
                strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    def my_block_request(self, *args, **kwargs):
        dfs = []
        kwargs0 = {**kwargs, **{'next': 0}}
        df = self.kiwoom.block_request(*args, **kwargs0)
        dfs.append(df)
        QTest.qWait(200)
        while self.kiwoom.tr_remained:
            kwargs2 = {**kwargs, **{'next': 2}}
            df = self.kiwoom.block_request(*args, **kwargs2)
            dfs.append(df)
            QTest.qWait(200)

        dfs = pd.concat(dfs)
        dfs = dfs[dfs.iloc[:, 0] != '']
        dfs = dfs.apply(self.convert_type)
        return dfs

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()

    # balance, item_list = myWindow.check_balance()
    # portfolio = myWindow.btn_file_load_clicked()

    app.exec_()
