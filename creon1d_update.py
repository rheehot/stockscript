# STOCK데이터베이스 안의 stock_1d테이블 하나에 몰아넣는 코드, 날짜 int형, 일별 데이여서 날짜시간 구별 없음

import win32com.client
from pywinauto import application
from datetime import datetime
import os
import sys
import pandas as pd
import pyodbc
import time
import numpy as np


###########################################################


Down_1d_data = True                 # 1일봉 다운 여부
Down_1m_data = True                 # 1분봉 다운 여부
Down_1w_data = True                 # 1주봉 다운 여부
Update_Modified_Price = True        # 수정주가 처리 여부
Down_index_data = True              # 지수 다운 여부
Down_market_type = True              # 마켓타입 다운 여부

START_DATE =    19800101             # 시작일
END_DATE =      int(datetime.today().strftime("%Y%m%d"))             # 종료일

weekDay = ['Mon','Tue',"Wed",'Thu','Fri','Sat','Sun']
###########################################################


class Creon:
    def __init__(self):
        self.obj_CpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
        self.obj_CpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        self.obj_StockChart = win32com.client.Dispatch("CpSysDib.StockChart")

        # 서버에 존재하는 종목코드 리스트
        self.sv_code_df = pd.DataFrame()
        self.sv_index_code_df = pd.DataFrame({'종목코드': ['U001','U201'], '종목명': ['KOSPI','KOSDAQ']},
                                       columns=('종목코드', '종목명'))

        # 데이터 기간용
        self.fromDate = START_DATE      # 굳이 fromDate를 만든 이유는 추가 데이터 업데이트시 새 날짜 기준부터 다운하기 위해서
        self.toDate = END_DATE
        self.isRecent = False

    ###### 로그인 관련 함수들
    def kill_client(self):
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')

    def connect(self, id_, pwd, pwdcert):
        if not self.connected():
            self.disconnect()
            self.kill_client()
            app = application.Application()
            app.start(
                'C:\CREON\STARTER\coStarter.exe /prj:cp /id:{id} /pwd:{pwd} /pwdcert:{pwdcert} /autostart'.format(
                    id=id_, pwd=pwd, pwdcert=pwdcert
                )
            )
        while not self.connected():
            time.sleep(1)
        return True

    def connected(self):
        b_connected = self.obj_CpCybos.IsConnect
        if b_connected == 0:
            return False
        return True

    def disconnect(self):
        if self.connected():
            self.obj_CpCybos.PlusDisconnect()

    ##########################################################################################


    def down_chart(self, code, search_type, date_from, date_to):

        self._wait()

        while(1):
            b_connected = self.obj_CpCybos.IsConnect
            if b_connected == 0:
                print("연결 실패 연결 재시도")
                time.sleep(10)
                self.connect('크레온아이디', '비번', '공인인증서비번')
            else:
                break

        # 0:날짜 1:시간hhmm 2:시가 3:고가 4:저가 5:종가 8:거래량 9:거래대금 10:누적체결매도수량
        # 11:누적체결매수수량 12:상장주식수 13:시가총액 14:외국인주문한도수량 15:외국인주문가능수량
        # 16:외국인현보유수량 17:외국인현보유비율 18:수정주가일자YYYYMMDD 19:수정주가비율
        # 20:기관순매수 21:기관누적순매수 22:등락주선 23:등락비율 24:예탁금 25:주식회전율 26:거래성립률
        down_type = 0
        if search_type == '1d':
            # 일봉 데이터
            list_field_key = [0, 2, 3, 4, 5, 8, 9, 10, 11, 13, 17, 18, 19, 20, 21]
            list_field_name = ['날짜', '시가', '고가', '저가', '종가', '거래량', '거래대금', '누적체결매도수량',
                           '누적체결매수수량', '시가총액', '외국인현보유비율','수정주가일자','수정주가비율',
                           '기관순매수','기관누적순매수']
            down_type = ord('D')

        elif search_type == '1w':
            # 주봉 데이터
            list_field_key = [0, 2, 3, 4, 5, 8, 9, 10, 11, 13, 17, 18, 19, 20, 21]
            list_field_name = ['날짜', '시가', '고가', '저가', '종가', '거래량', '거래대금', '누적체결매도수량',
                           '누적체결매수수량', '시가총액', '외국인현보유비율','수정주가일자','수정주가비율',
                           '기관순매수','기관누적순매수']
            down_type = ord('W')


        elif search_type == '1m':
            # 분봉 데이터
            list_field_key = [0, 1, 2, 3, 4, 5, 8, 9, 10, 11]
            list_field_name = ['날짜', '시간', '시가', '고가', '저가', '종가', '거래량', '거래대금', '누적매도수량',
                               '누적매수수량']
            down_type = ord('m')

        else:
            return None


        dict_chart = {name: [] for name in list_field_name}

        self.obj_StockChart.SetInputValue(0, code)
        self.obj_StockChart.SetInputValue(1, ord('2'))  # 2: 개수, 1: 기간
        #self.obj_StockChart.SetInputValue(2, date_to)  # 종료일
        #self.obj_StockChart.SetInputValue(3, date_from)  # 시작일
        self.obj_StockChart.SetInputValue(4, 200000)  # 요청 개수
        self.obj_StockChart.SetInputValue(5, list_field_key)  # 필드
        self.obj_StockChart.SetInputValue(6, down_type)  # D:일  W:주  M:월  m:분  T:틱
        self.obj_StockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.obj_StockChart.SetInputValue(10, ord('1'))  # 장종료시간외거래량만포함

        while True:
            self.obj_StockChart.BlockRequest()

            status = self.obj_StockChart.GetDibStatus()
            msg = self.obj_StockChart.GetDibMsg1()

            if status != 0:
                print("통신상태:", status, msg)
                return None

            cnt = self.obj_StockChart.GetHeaderValue(3)  # 수신개수

            for i in range(cnt):
                dict_item = {name: self.obj_StockChart.GetDataValue(pos, i) for pos, name in
                             zip(range(len(list_field_name)), list_field_name)}
                for k, v in dict_item.items():
                    dict_chart[k].append(v)

            if not self.obj_StockChart.Continue:
                break
            if int(dict_chart['날짜'][-1]) < int(date_from):
                break

            self._wait()

        stock = pd.DataFrame(dict_chart, columns=list_field_name)       # dataframe으로 만들기
        #print(date_from, '  ', date_to, stock)
        # 요청했던 만큼만 저장하기
        stock = stock[stock['날짜']>=date_from]


        # 중복으로 받은 행이 있다면 제거
        if search_type=='1d' or search_type=='1w':
            stock = stock.drop_duplicates('날짜', keep='first')
        elif search_type=='1m':
            stock = stock.drop_duplicates(['날짜','시간'], keep='first')

        # 누적거래대금 만들기
        if search_type == '1m':
            stock = self.ADD_Acc_Vol(stock)

        return stock

    # 누적거래대금 만드는 함수
    def ADD_Acc_Vol(self, data_table):
        data_table['누적거래대금'] = np.nan
        return_data = pd.DataFrame(columns=data_table.columns.tolist())  # 누적거래량 컬럼 생성

        dates = data_table['날짜'].sort_values(ascending=True).unique()

        for day in dates:
            set_data = data_table[data_table.날짜 == day]
            set_data = set_data.sort_values(by='시간')

            set_data.loc[:, '누적거래대금'] = set_data['거래대금'].cumsum()
            return_data = pd.concat([return_data, set_data])

        return return_data

    # 마켓에 해당하는 종목코드 리스트 반환하는 메소드
    def get_code_list(self, market):
        """
        :param market: 1:코스피, 2:코스닥, ...
        :return: market에 해당하는 코드 list
        """
        code_list = self.obj_CpCodeMgr.GetStockListByMarket(market)
        return code_list

    # 부구분코드를 반환하는 메소드
    def get_section_code(self, code):
        section_code = self.obj_CpCodeMgr.GetStockSectionKind(code)
        return section_code

    # 종목 코드를 받아 종목명을 반환하는 메소드
    def get_code_name(self, code):
        code_name = self.obj_CpCodeMgr.CodeToName(code)
        return code_name

    # def save_code(self):
    #     # 서버 종목 정보 가져와서 dataframe으로 저장
    #     sv_code_list = self.get_code_list(1) + self.get_code_list(2)
    #     sv_name_list = list(map(self.get_code_name, sv_code_list))
    #     self.sv_code_df = pd.DataFrame({'종목코드': sv_code_list, '종목명': sv_name_list},
    #                                    columns=('종목코드', '종목명'))

    def save_code(self):
        # 서버 종목 정보 가져와서 dataframe으로 저장
        sv_code_list = self.get_code_list(1)
        sv_name_list = list(map(self.get_code_name, sv_code_list))
        kospi = pd.DataFrame({'타입': 'KOSPI','종목코드': sv_code_list, '종목명': sv_name_list},
                                       columns=('타입','종목코드', '종목명'))

        sv_code_list = self.get_code_list(2)
        sv_name_list = list(map(self.get_code_name, sv_code_list))
        kosdaq = pd.DataFrame({'타입': 'KOSDAQ', '종목코드': sv_code_list, '종목명': sv_name_list},
                             columns=('타입', '종목코드', '종목명'))

        self.sv_code_df = pd.concat([kospi,kosdaq], ignore_index=True)
        self.sv_code_df = self.sv_code_df.sort_values(by='종목코드',ascending=True)

    def _wait(self):
        time_remained = self.obj_CpCybos.LimitRequestRemainTime
        cnt_remained = self.obj_CpCybos.GetLimitRemainCount(1)  # 0: 주문 관련, 1: 시세 요청 관련, 2: 실시간 요청 관련
        if cnt_remained <= 0:
            timeStart = time.time()
            while cnt_remained <= 0:
                time.sleep(time_remained / 1000)
                time_remained = self.obj_CpCybos.LimitRequestRemainTime
                cnt_remained = self.obj_CpCybos.GetLimitRemainCount(1)





# 날짜와 시간 체크
now = datetime.now()
week = datetime.today().weekday()
if weekDay[week] == 'Mon' or weekDay[week] =='Tue' or weekDay[week] =="Wed" or weekDay[week] =='Thu' or weekDay[week] =='Fri':
    if int(now.strftime('%H'))>=6 and int(now.strftime('%H'))<=16:
        print('다운 가능 시간이 아닙니다')
        #os.system('Pause')
        sys.exit()

creon = Creon()
creon.connect('아이디','비번','공인인증서비번')

creon.save_code() # 서버에서 주식코드 조회, 저장

# MSSQL 연결
# 1일봉
conn_1d = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=STOCK1d_2;UID=아이디;PWD=비번')
cursor_1d = conn_1d.cursor()
# 1분봉
conn_1m = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=STOCK1m_3;UID=아이디;PWD=비번')
cursor_1m = conn_1m.cursor()
# 1주봉
conn_1w = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=STOCK1w;UID=아이디;PWD=비번')
cursor_1w = conn_1w.cursor()
# 지수
conn_index = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=PRICE_INDEX_1d;UID=아이디;PWD=비번')
cursor_index = conn_index.cursor()
# 마켓타입
conn_market = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=MARKET_TYPE;UID=아이디;PWD=비번')
cursor_market = conn_market.cursor()



start_time = datetime.now()
print('시작: ', start_time)

cursor_1d.execute("select * from sysobjects where name = 'stock_1d'")  # 테이블 있나 확인
isTable = cursor_1d.fetchone()
if (isTable == None):  # 테이블 유무 확인하고 테이블 만들기
    cursor_1d.execute("""
    CREATE TABLE stock_1d(
        종목코드 CHAR(7) NOT NULL,
        날짜 INT NOT NULL,
        시가 FLOAT NULL,
        고가 FLOAT NULL,
        저가 FLOAT NULL,
        종가 FLOAT NULL,
        거래량 INT NULL,
        거래대금 MONEY NULL,
        누적매도수량 INT NULL,
        누적매수수량 INT NULL,
        시가총액 MONEY NULL,
        외국인현보유비율 FLOAT NULL,
        수정주가일자 INT NULL,
        수정주가비율 FLOAT NULL,
        기관순매수 MONEY NULL,
        기관누적순매수 MONEY NULL
        CONSTRAINT [PK_stock_1d] PRIMARY KEY CLUSTERED 
        (
            [종목코드] ASC, 
            [날짜] ASC
        )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
    ) ON [PRIMARY]
    """)
    conn_1d.commit()
    creon.isRecent = False
    creon.fromDate = START_DATE


# 여기는 1d 데이터 처리하는곳, 만약 수정주가 처리할게 있다면 데이터 전부 지우고 받기 시작한다
for i in range(len(creon.sv_code_df)):

    if Down_1d_data != True:
        break

    print('[일봉] ',creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i],' 받는중... ', i,'/', len(creon.sv_code_df), datetime.now())

    cursor_1d.execute("select TOP(1) 날짜, 종가 from stock_1d where 종목코드=? ORDER BY 날짜 DESC", creon.sv_code_df['종목코드'][i])
    lastDay = cursor_1d.fetchone()


    if (lastDay == None):
        creon.fromDate = START_DATE
        creon.isRecent = False

    elif (int(lastDay[0]) >= int(creon.toDate)) :
        continue

    else:
        creon.fromDate = lastDay[0]


    stock_data = creon.down_chart(creon.sv_code_df['종목코드'][i],'1d', creon.fromDate, creon.toDate)
    stock_data['종목코드'] = creon.sv_code_df['종목코드'][i]
    stock_data = stock_data.sort_values(by='날짜',ascending=False)  # 날짜기준 내림차순정렬
    stock_data = stock_data.reset_index(drop=True)

    # 넣기전에 덜받은 마지막날 데이터는 지워준다
    if (lastDay != None):
        last_day_data = stock_data[stock_data['날짜']==lastDay[0]] # 기존 DB에 있던 마지막날 데이터 가져오기
        last_day_data = last_day_data.reset_index(drop=True)
        if(last_day_data['종가'][0] == lastDay[1]):
            stock_data = stock_data[stock_data['날짜'] > lastDay[0]]
            #print(stock_data)
        else:   # 수정주가 있으면 전체 지움
            print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 수정주가 발견, 전부 지우고 새로 MSSQL에 넣는중', datetime.now())
            stock_data = creon.down_chart(creon.sv_code_df['종목코드'][i], '1d', START_DATE, creon.toDate)
            cursor_1d.execute("delete from stock_1d where 종목코드=?", creon.sv_code_df['종목코드'][i])
            conn_1d.commit()

    if stock_data.empty:
        print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 데이터 조회결과 없어서 넘김')
        continue

    print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 데이터 MSSQL에 넣는중', datetime.now())

    stock_cnt = len(stock_data)
    for j in range(stock_cnt):
        cursor_1d.execute("INSERT INTO stock_1d VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                   (str(creon.sv_code_df['종목코드'][i]),
                    int(stock_data['날짜'][j]),
                    float(stock_data['시가'][j]),
                   float(stock_data['고가'][j]),
                   float(stock_data['저가'][j]),
                   float(stock_data['종가'][j]),
                   int(stock_data['거래량'][j]),
                   int((stock_data['거래대금'][j])),
                   int(stock_data['누적체결매도수량'][j]),
                   int(stock_data['누적체결매수수량'][j]),
                    int(stock_data['시가총액'][j]),
                    float(stock_data['외국인현보유비율'][j]),
                    int(stock_data['수정주가일자'][j]),
                    float(stock_data['수정주가비율'][j]),
                    int(stock_data['기관순매수'][j]),
                    int(stock_data['기관누적순매수'][j])))
    conn_1d.commit()



# 여기는 1m 데이터 처리하는곳, 일단 전부 다운받고 일봉과 비교해서 수정주가 처리해줌
for i in range(len(creon.sv_code_df)):

    if Down_1m_data != True:
        break

    print('[1분봉] ',creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i],' 받는중... ', i,'/', len(creon.sv_code_df), datetime.now())

    tablename = creon.sv_code_df['종목코드'][i]  # 테이블을 만들지 말지 결정하기 위해서 설정

    cursor_1m.execute("select * from sysobjects where name = '" + tablename + "'")  # 테이블 있나 확인
    isTable = cursor_1m.fetchone()
    if (isTable == None):  # 테이블 유무 확인하고 테이블 만들기
        cursor_1m.execute("""
        CREATE TABLE """ + tablename + """(
            종목코드 CHAR(7) NOT NULL,
            날짜 INT NOT NULL,
            시간 INT NOT NULL,
            시가 FLOAT NULL,
            고가 FLOAT NULL,
            저가 FLOAT NULL,
            종가 FLOAT NULL,
            거래량 BIGINT NULL,
            거래대금 MONEY NULL,
            누적매도수량 BIGINT NULL,
            누적매수수량 BIGINT NULL,
            누적거래대금 BIGINT NULL
            CONSTRAINT ["""+'PK_'+tablename+"""] PRIMARY KEY CLUSTERED
            (
                [종목코드] ASC,
                [날짜] ASC,
                [시간] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
        ) ON [PRIMARY]
        """)
        conn_1m.commit()
        creon.isRecent = False
        creon.fromDate = START_DATE

    else:  # 테이블이 있다면 테이블에서 최근 날짜 가져오기
        cursor_1m.execute("select TOP(1) 날짜 from " + tablename + " ORDER BY 날짜 DESC, 시간 DESC")
        lastDay = cursor_1m.fetchone()

        if (lastDay == None):
            creon.fromDate = START_DATE
            print(creon.sv_code_df['종목명'][i], '테이블이 비어있음')
            creon.isRecent = False

        elif (int(lastDay[0]) >= int(creon.toDate)):
            continue

        else:
            creon.fromDate = lastDay[0]



    stock_data = creon.down_chart(creon.sv_code_df['종목코드'][i],'1m', creon.fromDate, creon.toDate)
    stock_data = stock_data.sort_values(by=['날짜', '시간'],ascending=False)  # 날짜기준 내림차순정렬
    stock_data = stock_data.reset_index(drop=True)

    # 넣기전에 마지막날 데이터는 지워준다
    if (lastDay != None):
        stock_data = stock_data[stock_data['날짜'] > lastDay[0]]

    stock_data['종목코드'] = creon.sv_code_df['종목코드'][i]

    if stock_data.empty:
        print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 데이터 조회결과 없어서 넘김')
        continue

    print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 데이터 MSSQL에 넣는중', datetime.now())



    stock_cnt = len(stock_data)
    for j in range(stock_cnt):
        cursor_1m.execute("INSERT INTO " + tablename + " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                       (str(creon.sv_code_df['종목코드'][i]),
                        int(stock_data['날짜'][j]),
                        int(stock_data['시간'][j]),
                        float(stock_data['시가'][j]),
                        float(stock_data['고가'][j]),
                        float(stock_data['저가'][j]),
                        float(stock_data['종가'][j]),
                        int(stock_data['거래량'][j]),
                        int((stock_data['거래대금'][j])),
                        int(stock_data['누적매도수량'][j]),
                        int(stock_data['누적매수수량'][j]),
                        int(stock_data['누적거래대금'][j])))

    conn_1m.commit()



# 여기는 1w 데이터 처리하는곳, 만약 수정주가 처리할게 있다면 데이터 전부 지우고 받기 시작한다

cursor_1w.execute("select * from sysobjects where name = 'stock_1w'")  # 테이블 있나 확인
isTable = cursor_1w.fetchone()
if (isTable == None):  # 테이블 유무 확인하고 테이블 만들기
    cursor_1w.execute("""
    CREATE TABLE stock_1w(
        종목코드 CHAR(7) NOT NULL,
        날짜 INT NOT NULL,
        시가 FLOAT NULL,
        고가 FLOAT NULL,
        저가 FLOAT NULL,
        종가 FLOAT NULL,
        거래량 INT NULL,
        거래대금 MONEY NULL,
        누적매도수량 INT NULL,
        누적매수수량 INT NULL,
        시가총액 MONEY NULL,
        외국인현보유비율 FLOAT NULL,
        수정주가일자 INT NULL,
        수정주가비율 FLOAT NULL,
        기관순매수 MONEY NULL,
        기관누적순매수 MONEY NULL
        CONSTRAINT [PK_stock_1w] PRIMARY KEY CLUSTERED 
        (
            [종목코드] ASC, 
            [날짜] ASC
        )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
    ) ON [PRIMARY]
    """)
    conn_1w.commit()
    creon.isRecent = False
    creon.fromDate = START_DATE

for i in range(len(creon.sv_code_df)):

    if Down_1w_data != True:
        break

    if weekDay[week] == 'Mon' or weekDay[week] == 'Tue' or weekDay[week] == "Wed" or weekDay[week] == 'Thu':
        break;
    elif  weekDay[week] == 'Fri' and int(now.strftime('%H'))<=16:
        break;

    print('[주봉] ',creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i],' 받는중... ', i,'/', len(creon.sv_code_df), datetime.now())

    cursor_1w.execute("select TOP(1) 날짜, 종가 from stock_1w where 종목코드=? ORDER BY 날짜 DESC", creon.sv_code_df['종목코드'][i])
    lastDay = cursor_1w.fetchone()


    if (lastDay == None):
        creon.fromDate = START_DATE
        creon.isRecent = False

    elif (int(lastDay[0]) >= int(creon.toDate)) :
        continue

    else:
        creon.fromDate = lastDay[0]


    stock_data = creon.down_chart(creon.sv_code_df['종목코드'][i],'1w', creon.fromDate, creon.toDate)
    stock_data['종목코드'] = creon.sv_code_df['종목코드'][i]
    stock_data = stock_data.sort_values(by='날짜',ascending=False)  # 날짜기준 내림차순정렬
    stock_data = stock_data.reset_index(drop=True)

    # 넣기전에 덜받은 마지막날 데이터는 지워준다
    if (lastDay != None):
        last_day_data = stock_data[stock_data['날짜']==lastDay[0]] # 기존 DB에 있던 마지막날 데이터 가져오기
        last_day_data = last_day_data.reset_index(drop=True)
        print(last_day_data)
        if(last_day_data['종가'][0] == lastDay[1]):
            stock_data = stock_data[stock_data['날짜'] > lastDay[0]]
            #print(stock_data)
        else:   # 수정주가 있으면 전체 지움
            print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 수정주가 발견, 전부 지우고 새로 MSSQL에 넣는중', datetime.now())
            stock_data = creon.down_chart(creon.sv_code_df['종목코드'][i], '1d', START_DATE, creon.toDate)
            cursor_1w.execute("delete from stock_1w where 종목코드=?", creon.sv_code_df['종목코드'][i])
            conn_1d.commit()

    if stock_data.empty:
        print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 데이터 조회결과 없어서 넘김')
        continue

    print(creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 데이터 MSSQL에 넣는중', datetime.now())

    stock_cnt = len(stock_data)
    for j in range(stock_cnt):
        cursor_1w.execute("INSERT INTO stock_1w VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                   (str(creon.sv_code_df['종목코드'][i]),
                    int(stock_data['날짜'][j]),
                    float(stock_data['시가'][j]),
                   float(stock_data['고가'][j]),
                   float(stock_data['저가'][j]),
                   float(stock_data['종가'][j]),
                   int(stock_data['거래량'][j]),
                   int((stock_data['거래대금'][j])),
                   int(stock_data['누적체결매도수량'][j]),
                   int(stock_data['누적체결매수수량'][j]),
                    int(stock_data['시가총액'][j]),
                    float(stock_data['외국인현보유비율'][j]),
                    int(stock_data['수정주가일자'][j]),
                    float(stock_data['수정주가비율'][j]),
                    int(stock_data['기관순매수'][j]),
                    int(stock_data['기관누적순매수'][j])))
    conn_1w.commit()


# 여기는 1분봉 수정주가 처리하는곳
for i in range(len(creon.sv_code_df)):

    if Update_Modified_Price != True:
        break

    print('[수정주가] ',creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i],' 수정주가 처리중... ', i,'/', len(creon.sv_code_df), datetime.now())

    tablename = creon.sv_code_df['종목코드'][i]  # 테이블을 만들지 말지 결정하기 위해서 설정
    cursor_1m.execute("""
        update l1
            set l1.시가 = ROUND(l1.시가 * l2.ratio,0)
            , l1.고가 = ROUND(l1.고가 * l2.ratio,0)
            , l1.저가 = ROUND(l1.저가 * l2.ratio,0)
            , l1.종가 = ROUND(l1.종가 * l2.ratio,0)
        from [STOCK1m_3].dbo."""+tablename+""" l1 join (
            select 종목코드, 날짜, 시가 / 분봉시가 * 1.0 ratio from
            (
                select 종목코드, 날짜, 시가
                , (SELECT TOP (1) 시가 FROM  [STOCK1m_3].dbo."""+tablename+""" WHERE 종목코드 = d.종목코드 AND 날짜 = d.날짜 ORDER BY 시간 ) 분봉시가
                from [STOCK1d_2].dbo.stock_1d d
            ) a where ABS(분봉시가 - 시가) > 1
        ) l2 on l1.종목코드 = l2.종목코드 and l1.날짜 = l2.날짜
    """)
    conn_1m.commit()



# 코스피, 코스닥지수 입력
for i in range(len(creon.sv_index_code_df)):

    if Down_index_data != True:
        break

    print(creon.sv_index_code_df['종목코드'][i], creon.sv_index_code_df['종목명'][i],' 받는중... ', i,'/', len(creon.sv_index_code_df), datetime.now())

    tablename = creon.sv_index_code_df['종목명'][i]

    cursor_index.execute("select * from sysobjects where name = ?", tablename)  # 테이블 있나 확인
    isTable = cursor_index.fetchone()
    if (isTable == None):  # 테이블 유무 확인하고 테이블 만들기
        cursor_index.execute("""
        CREATE TABLE """+ tablename +"""(
            종목코드 CHAR(7) NOT NULL,
            날짜 INT NOT NULL,
            시가 FLOAT NULL,
            고가 FLOAT NULL,
            저가 FLOAT NULL,
            종가 FLOAT NULL,
            거래량 BIGINT NULL,
            거래대금 BIGINT NULL,
            시가총액 BIGINT NULL
            CONSTRAINT [PK_"""+tablename+"""] PRIMARY KEY CLUSTERED 
            (
                [종목코드] ASC, 
                [날짜] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
        ) ON [PRIMARY]
        """)
        conn_index.commit()
        creon.isRecent = False
        creon.fromDate = START_DATE

    cursor_index.execute("select TOP(1) 날짜 from "+tablename+" where 종목코드=? ORDER BY 날짜 DESC", creon.sv_index_code_df['종목코드'][i])
    lastDay = cursor_index.fetchone()

    if (lastDay == None):
        creon.fromDate = START_DATE
        creon.isRecent = False

    elif (int(lastDay[0]) >= int(creon.toDate)) :
        continue

    else:
        creon.fromDate = lastDay[0]


    stock_data = creon.down_chart(creon.sv_index_code_df['종목코드'][i], '1d', creon.fromDate, creon.toDate)
    stock_data = stock_data.sort_values(by='날짜',ascending=False)  # 날짜기준 내림차순정렬
    stock_data['종목코드'] = creon.sv_index_code_df['종목코드'][i]
    stock_data = stock_data.reset_index(drop=True)

    # 넣기전에 덜받은 마지막날 데이터는 지워준다
    stock_data = stock_data[stock_data['날짜'] > int(creon.fromDate)]

    if stock_data.empty:
        print(creon.sv_index_code_df['종목코드'][i], creon.sv_index_code_df['종목명'][i], ' 데이터 조회결과 없어서 넘김')
        continue

    print(creon.sv_index_code_df['종목코드'][i], creon.sv_index_code_df['종목명'][i], ' 데이터 MSSQL에 넣는중', datetime.now())

    stock_cnt = len(stock_data)
    for j in range(stock_cnt):
        cursor_index.execute("INSERT INTO "+tablename+" VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                   (str(creon.sv_index_code_df['종목코드'][i]),
                    int(stock_data['날짜'][j]),
                    float(stock_data['시가'][j]),
                   float(stock_data['고가'][j]),
                   float(stock_data['저가'][j]),
                   float(stock_data['종가'][j]),
                   int(stock_data['거래량'][j]),
                   int((stock_data['거래대금'][j])),
                    int(stock_data['시가총액'][j])))
    conn_index.commit()




# 마켓 타입 다운로드
cursor_market.execute("select * from sysobjects where name = 'market_type'")  # 테이블 있나 확인
isTable = cursor_market.fetchone()
if (isTable == None):  # 테이블 유무 확인하고 테이블 만들기
    cursor_market.execute("""
    CREATE TABLE market_type(
        타입 CHAR(6) NOT NULL,
        종목코드 CHAR(7) NOT NULL,
        종목명 NVARCHAR(50) NOT NULL
        CONSTRAINT [PK_TYPE] PRIMARY KEY CLUSTERED 
        (
            [종목코드] ASC, 
            [종목명] ASC
        )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
    ) ON [PRIMARY]
    """)
    conn_market.commit()
    creon.isRecent = False
    creon.fromDate = START_DATE

# 여기서부턴 for문으로 모든 종목 조회되게 고쳐야할 부분
for i in range(len(creon.sv_code_df)):

    if Down_market_type != True:
        break

    cursor_market.execute("select * from market_type where 종목코드 = ?",creon.sv_code_df['종목코드'][i])  # 테이블 있나 확인
    isType = cursor_market.fetchone()

    if isType == None:
        print('마켓타입', creon.sv_code_df['종목코드'][i], creon.sv_code_df['종목명'][i], ' 넣는중..')
        cursor_market.execute("INSERT INTO market_type VALUES (?, ?, ?)",
                   (str(creon.sv_code_df['타입'][i]),
                    str(creon.sv_code_df['종목코드'][i]),
                    creon.sv_code_df['종목명'][i]))
        conn_market.commit()


end_time = datetime.now()
print('끝: ', end_time)
print('총 걸린시간: ', end_time-start_time)
print('end')