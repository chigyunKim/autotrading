import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import openpyxl
#from slacker import Slacker
import time, calendar


def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)


def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    #slack.chat.post_message('#etf-algo-trading', strbuf)


# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')


def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('check_creon_system() : admin user -> FAILED')
        return False

    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        printlog('check_creon_system() : connect to server -> FAILED')
        return False

    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True


def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
    cpOhlc.SetInputValue(0, code)  # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))  # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)  # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 0:날짜, 2~5:OHLC, 8:거래량
    cpOhlc.SetInputValue(6, ord('D'))  # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))  # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)  # 3:수신개수
    columns = ['open', 'high', 'low', 'close', 'volume']
    index = []
    rows = []
    for i in range(count):
        index.append(cpOhlc.GetDataValue(0, i))
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
                     cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i), cpOhlc.GetDataValue(5,i)])
    df = pd.DataFrame(rows, columns=columns, index=index)
    return df


if __name__ == '__main__':
    try:
        etfuniverse = pd.read_excel(r'C:\Users\chigy\Desktop\superETF\ETF_universe.xlsx')
        etfcode = etfuniverse['code']
        symbol_list = []

        for code in range(0,len(etfcode)):
            if len(str(etfcode[code])) == 6 :
                symbol_list.append('A'+ str(etfcode[code]))
            else :
                symbol_list.append('A0' + str(etfcode[code]))

        printlog('check_creon_system() :', check_creon_system())  # 크레온 접속 점검

        tmpDict = {}

        for i in range(0,len(symbol_list)) : #대상 ETF 중에서 20일 평균거래량이 가장 큰 ETF의 code 구하기
            globals()['df'+ str(symbol_list[i])] = get_ohlc(symbol_list[i], 20)
            globals()['df' + str(symbol_list[i])].sort_index(ascending=True, inplace=True)  # 과거에서 현재순으로 정렬
            globals()['df' + str(symbol_list[i])]['volavg20d'] = globals()['df' + str(symbol_list[i])]['volume'].rolling(20).mean()  # 20일 평균거래량 계산
            #print(globals()['df'+str(symbol_list[i])])
            print(str(symbol_list[i]))
            print(globals()['df'+str(symbol_list[i])].loc[globals()['df'+str(symbol_list[i])].index[-1]]['volavg20d'])
            print('')
            tmpDict[str(symbol_list[i])] = globals()['df'+str(symbol_list[i])].loc[globals()['df'+str(symbol_list[i])].index[-1]]['volavg20d']
            time.sleep(1)

        targetETF = max(tmpDict, key = tmpDict.get)
        print(tmpDict)
        print(targetETF)

    except Exception as ex:
        dbgout('`main -> exception! ' + str(ex) + '`')