import sys
import os

from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from beautifultable import BeautifulTable

from autokiwoom.consts import *
from autokiwoom.message import outputmsg

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("kiwoom class")

        # 이벤트 루프 변수
        self.login_event_loop = QEventLoop()            # 로그인 담당 이벤트 루프
        self.ac_loop = QEventLoop()                     # 계좌관련 이벤트 루프

        # 계좌 관련 변수
        self.acno = None                                # 계좌번호
        self.tot_buy_bal = 0                            # 총 매입 금액
        self.tot_eval_bal = 0                           # 총 평가 금액
        self.tot_eval_pl_bal = 0                        # 총 평가손익 금액
        self.tot_pl_rt = 0                              # 총 수익률
        self.ac_stock_dict = {}                         # 계좌 보유 주식 Dictionary
        self.incomplete_order_dict = {}                 # 미체결주문 Dictionary

        # 예수금 관련 변수
        self.total_bal = 0                              # 예수금
        self.withdraw_bal = 0                           # 출금가능 금액
        self.order_bal = 0                              # 주만가능 금액

        # 화면번호
        self.scrno = None

        # 초기 작업
        self.create_instance()
        self.event_collection()                         # 이벤트와 슬롯을 메모리에 먼저 생성.
        self.login()
        #input()
        self.get_account_info()                         # 계좌 번호만 얻어오기
        self.get_deposit_info()                         # 예수금 관련된 정보 얻어오기
        self.get_ac_eval_bal()                          # 계좌 평가잔고 내역 얻어오기
        self.get_incomplete_order()                     # 미체결 주문내역 얻어오기

        self.menu()

    # COM 오브젝트 생성
    def create_instance(self):
        self.setControl(KIWOOM_API_PROGID)              # 레즈스트리에 저장된 키움 openAPI 모듈 불러오기

    def event_collection(self):
        self.OnEventConnect.connect(self.login_slot)    # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.tr_slot)      # 트랙잭션 요청 관련 이벤트

    def login(self):
        self.dynamicCall("CommConnect()")               # 시그널 함수 호출.
        self.login_event_loop.exec_()

    def login_slot(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            os.system('cls')
            print("에러 내용 : ", outputmsg(err_code)[1])
            sys.exit(0)
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCLIST")
        self.acno = account_list.split(";")[0]
        print(self.acno)

    def menu(self):
        sel = ""
        while True:
            os.system('cls')
            print("========================================================")
            print("1. 현재 로그인 상태 확인")
            print("2. 사용자 정보 조회")
            print("3. 예수금 조회")
            print("4. 계좌 평가잔고 조회")
            print("5. 미체결내역 조회")
            print("========================================================")
            print("0. 프로그램 종료")
            print("========================================================")

            sel = input("=> ")

            if sel == "0":
                sys.exit(0)
            elif sel == "1":
                self.print_login_connect_state()
            elif sel == "2":
                self.print_my_info()
            elif sel == "3":
                self.print_deposit_info()
            elif sel == "4":
                self.print_ac_eval_bal_info()
            elif sel == "5":
                self.print_incomplete_order()
            else:
                print(sel, "은 없는 구분값입니다. 다시 입력해주세요.")
                input()

    def print_login_connect_state(self):
        os.system('cls')
        isLogin = self.dynamicCall("GetConnectState()")
        if isLogin == 1:
            print("\n현재 계정은 로그인 상태입니다.\n")
        else:
            print("\n현재 계정은 로그아웃 상태입니다.\n")
        input()

    def print_my_info(self):
        os.system('cls')
        user_name   = self.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        user_id     = self.dynamicCall("GetLoginInfo(QString)", "USER_ID")
        ac_cnt      = self.dynamicCall("GetLoginInfo(QString)", "ACCOUNT_CNT")
        
        print(f"\n이름              : {user_name}")
        print(f"ID                : {user_id}")
        print(f"보유 계좌 수      : {ac_cnt}")
        print(f"1번째 계좌번호    : {self.acno}\n")
        input()

    def print_deposit_info(self):
        os.system('cls')
        print(f"\n예수금              : {self.total_bal} 원")
        print(f"출금 가능 금액      : {self.withdraw_bal} 원")
        print(f"주문 가능 금액      : {self.order_bal} 원\n")
        input()

    def print_ac_eval_bal_info(self):
        os.system('cls')
        print(f"\n총 매입 금액         : {self.tot_buy_bal} 원")
        print(f"총 평가 금액         : {self.tot_eval_bal} 원")
        print(f"총 평가손익 금액     : {self.tot_eval_pl_bal} 원")
        print(f"총 수익률            : {self.tot_pl_rt} %\n")

        table = self.make_table(TR_NAME_DEPOSIT_EVAL_BAL)
        print("\n---------------------------------------------------------\n")
        print("계좌 보유 종목별 평가잔고 현황")
        print(f"보유 종목 수         : {len(self.ac_stock_dict)} 개")
        print(table)
        input()    

    def print_incomplete_order(self):
        os.system('cls')
        print()
        
        if len(self.incomplete_order_dict) == 0:
            print("미체결 내역이 없습니다.")
        else:
            table = self.make_table(TR_NAME_INCOMPLETE_ORDER)
            print(table)
        
        input()    

    def make_table(self, sRQName):
        table = BeautifulTable()
        table = BeautifulTable(maxwidth=150)

        if sRQName == TR_NAME_DEPOSIT_EVAL_BAL:
            for stock_code in self.ac_stock_dict:
                stock = self.ac_stock_dict[stock_code]
                stockList = []
                for key in stock:
                    output = None

                    if key == "종목명":
                        output = stock[key]
                    elif key == "수익률(%)":
                        output = str(stock[key]) + "%"
                    elif key == "보유수량" or key == "매매가능수량":
                        output = str(stock[key]) + "개"
                    else:
                        output = str(stock[key]) + "원"
                    
                    stockList.append(output)
                table.rows.append(stockList)
            
            table.columns.header = ["종목명", "평가손익", "수익률", "매입가", "보유수량", "매매가능수량", "현재가"]
            #table.rows.sort('종목명')
        elif sRQName == TR_NAME_INCOMPLETE_ORDER:
            for stock_order_number in self.incomplete_order_dict:
                stock = self.ac_stock_dict[stock_order_number]
                stockList = [stock_order_number]
                for key in stock:
                    output = None

                    if key == "주문가격" or key == "현재가":
                        output = str(stock[key]) + "원"
                    elif "량" in key:
                        output = str(stock[key]) + "개"
                    elif key == "종목코드":
                        continue
                    else:
                        output = stock[key]
                    
                    stockList.append(output)
                table.rows.append(stockList)
            
            table.columns.header = ["주문번호", "종목코드", "종목명", "주문구분", "주문가격", "주문수량", "미체결수량", "체결량", "현재가", "주문상태"]
            table.rows.sort('주문번호')
        
        return table

    def get_deposit_info(self, nPrevNext=0):
        self.scrno = SCRNO_DEPOSIT_INFO
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.acno)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", " ")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")

        self.dynamicCall("CommRqData(QString, QString, int, QString)", TR_NAME_DEPOSIT_DETAIL_INFO, TR_CODE_DEPOSIT_DETAIL_INFO, nPrevNext, self.scrno)

        self.ac_loop.exec_()

    def get_ac_eval_bal(self, nPrevNext=0):
        self.scrno = SCRNO_DEPOSIT_INFO

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.acno)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", " ")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")

        self.dynamicCall("CommRqData(QString, QString, int, QString)", TR_NAME_DEPOSIT_EVAL_BAL, TR_CODE_DEPOSIT_EVAL_BAL, nPrevNext, self.scrno)

        # 페이징 조회 시 이벤트루프 수행상태 확인 필요.
        if not self.ac_loop.isRunning():
            self.ac_loop.exec_()

    def get_incomplete_order(self, nPrevNext=0):
        self.scrno = SCRNO_DEPOSIT_INFO

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.acno)
        self.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")

        self.dynamicCall("CommRqData(QString, QString, int, QString)", TR_NAME_INCOMPLETE_ORDER, TR_CODE_INCOMPLETE_ORDER, nPrevNext, self.scrno)
        print("HERE")
        # 페이징 조회 시 이벤트루프 수행상태 확인 필요.
        if not self.ac_loop.isRunning():
            self.ac_loop.exec_()

    def tr_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == TR_NAME_DEPOSIT_DETAIL_INFO:              # 예수금상세현황요청
            total_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.total_bal = int(total_bal)

            withdraw_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.withdraw_bal = int(withdraw_bal)

            order_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "주문가능금액")
            self.order_bal = int(order_bal)

            self.cancel_screen_number(self.scrno)               # 사용한 화면번호 끊어내기. (통신 종료처리를 위함)
            
            self.ac_loop.exit()                        # 이벤트 루프 종료
        elif sRQName == TR_NAME_DEPOSIT_EVAL_BAL:               # 계좌평가잔고내역요청
            tot_buy_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
            self.tot_buy_bal = int(tot_buy_bal)

            tot_eval_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가금액")
            self.tot_eval_bal = int(tot_eval_bal)

            tot_eval_pl_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.tot_eval_pl_bal = int(tot_eval_pl_bal)

            tot_pl_rt = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.tot_pl_rt = float(tot_pl_rt)

            # 반복부 처리
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(cnt):
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                stock_code = stock_code.strip()[1:]

                stock_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_name = stock_name.strip()

                stock_eval_pl = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "평가손익")
                stock_eval_pl = int(stock_eval_pl)

                stock_pl_rt = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                stock_pl_rt = float(stock_pl_rt)

                stock_buy_bal = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                stock_buy_bal = int(stock_buy_bal)

                stock_qty = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                stock_qty = int(stock_qty)

                stock_trade_qty = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                stock_trade_qty = int(stock_trade_qty)

                stock_cur_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                stock_cur_price = int(stock_cur_price)

                if not stock_code in self.ac_stock_dict:
                    self.ac_stock_dict[stock_code] = {}

                self.ac_stock_dict[stock_code].update({'종목명'         : stock_name})
                self.ac_stock_dict[stock_code].update({'평가손익'       : stock_eval_pl})
                self.ac_stock_dict[stock_code].update({'수익률(%)'      : stock_pl_rt})
                self.ac_stock_dict[stock_code].update({'매입가'         : stock_buy_bal})
                self.ac_stock_dict[stock_code].update({'보유수량'       : stock_qty})
                self.ac_stock_dict[stock_code].update({'매매가능수량'   : stock_trade_qty})
                self.ac_stock_dict[stock_code].update({'현재가'         : stock_cur_price})


            if sPrevNext == "2":                                # 데이터가 더 있는 경우는 sPrevNext가 2로 반환됨
                self.get_ac_eval_bal("2")
            else:
                self.cancel_screen_number(self.scrno)           # 사용한 화면번호 끊어내기. (통신 종료처리를 위함)
                self.ac_loop.exit()                # 이벤트 루프 종료
        elif sRQName == TR_NAME_INCOMPLETE_ORDER:               # 실시간미체결요청
            # 반복부 처리
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(cnt):
                stock_code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                stock_code = stock_code.strip()

                stock_order_number = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                stock_order_number = int(stock_order_number)

                stock_name = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_name = stock_name.strip()

                stock_order_type = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분")
                stock_order_type = stock_order_type.strip().lstrip('+').lstrip('-')

                stock_order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                stock_order_price = int(stock_order_price)

                stock_order_qty = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                stock_order_qty = int(stock_order_qty)

                stock_incomplete_order_qty = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                stock_incomplete_order_qty = int(stock_incomplete_order_qty)

                stock_complete_order_qty = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")
                stock_complete_order_qty = int(stock_complete_order_qty)

                stock_cur_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                stock_cur_price = int(stock_cur_price.strip().lstrip('+').lstrip('-'))

                stock_order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")
                stock_order_status = stock_order_status.strip()

                if not stock_code in self.ac_stock_dict:
                    self.ac_stock_dict[stock_code] = {}

                self.ac_stock_dict[stock_order_number].update({"종목코드"       : stock_code})
                self.ac_stock_dict[stock_order_number].update({"종목명"         : stock_name})
                self.ac_stock_dict[stock_order_number].update({"주문구분"       : stock_order_type})
                self.ac_stock_dict[stock_order_number].update({"주문가격"       : stock_order_price})
                self.ac_stock_dict[stock_order_number].update({"주문수량"       : stock_order_qty})
                self.ac_stock_dict[stock_order_number].update({"미체결수량"     : stock_incomplete_order_qty})
                self.ac_stock_dict[stock_order_number].update({"체결량"         : stock_complete_order_qty})
                self.ac_stock_dict[stock_order_number].update({"현재가"         : stock_cur_price})
                self.ac_stock_dict[stock_order_number].update({"주문상태"       : stock_order_status})


            if sPrevNext == "2":                                # 데이터가 더 있는 경우는 sPrevNext가 2로 반환됨
                self.get_incomplete_order("2")
            else:
                self.cancel_screen_number(self.scrno)           # 사용한 화면번호 끊어내기. (통신 종료처리를 위함)
                self.ac_loop.exit()                # 이벤트 루프 종료
            
    def cancel_screen_number(self, sScrNo):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)
