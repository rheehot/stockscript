# 일음님 당일단타 연구용 백테


import pandas as pd
from datetime import datetime, timedelta
import pyodbc
import numpy as np
import math

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib
# matplotlib.use('TkAgg')
fig, ax = plt.subplots(figsize=(8, 8))
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m/%d/%Y'))
plt.gca().xaxis.set_major_locator(mdates.MonthLocator(interval=6))
#ax.set_title('breakthrough')
#
# def plot_equity():
#     fig.autofmt_xdate()
#     plt.legend(['equity curve'], loc='upper left')
#     plt.show()

class Strategy_Option:
    def __init__(self):
        # 백테스트기간
        self.Start_date = 20190101
        self.End_date =  20191125

        # 초기자금
        self.Start_Money = 10000000
        self.Money = np.nan
        self.Equity = np.nan

        # 비용
        self.Fee = 0.00015
        self.Tax = 0.0025

        # 보유 종목수
        self.maximum_buy = 20
        self.buy_ratio = 1/self.maximum_buy

        # 목표가 손절가
        self.target_price = 1.05
        self.loss_price = 0.96

        # 타입컷
        self.buy_timecut = 9999
        self.sell_timecut = 1520
        # 상대적 타임컷
        self.relative_timecut = 30
        # 거래대금 기준
        self.tradeAccMoney = 10000000000

        # 검색 고가 가격 기준 (%로)
        ################################# 건드렸으면 일봉 검색기준도 같이 수정해라 ########################
        self.search_VI = 1.095
        self.search_price = 1.15

        # 시가갭기준 ( 시가갭 x 이하인 종목만 매수)
        self.open_gap_price = 1.03

        # 매수후보군
        self.trade_columns = ['종목코드', '종목명', '매수일', '매수시간', '매수가격', '매도일', '매도시간', '매도가격','수익률', '수익금']
        self.trade_list = pd.DataFrame(columns=self.trade_columns)

        # 현재 보유종목
        self.port_dict = {}
        # 매도예정종목
        self.sell_list = []

    def set_equity(self, day, current_money):
        self.Money[day] = current_money
        self.Equity[day] = current_money
        stock_val = 0
        for codename, port in self.port_dict.items():
            stock_val += port['buyPrice'] * port['amount']
        self.Equity[day] += stock_val

    def get_cagr(self):
        total_ret = (self.Equity[self.Equity.index[-1]] / self.Equity[self.Equity.index[0]] )
        total_day = datetime.strptime(str(self.Equity.index[-1]),'%Y%m%d')-datetime.strptime(str(self.Equity.index[0]),'%Y%m%d')
        total_day = total_day.days

        total_year = float(total_day) / 365

        return (total_ret ** (1 / total_year)) - 1 # ** 는 거듭제곱의 의미

    def get_mdd(self):
        dds = [0]
        max_equity = self.Equity.iloc[0]
        for i in self.Equity.index: #range(0, len(self.Equity)):
            if max_equity < self.Equity[i]:
                max_equity = self.Equity[i]
            dd = -(max_equity - self.Equity[i]) / max_equity
            dds.append(dd)
        return min(dds)

    def calcBuyPrice(self, high, low, type):
        price = low+(high-low)/8*5

        if price<1000:
            price = math.trunc(price)
        elif 1000 <= price < 5000:
            price = math.trunc(price/5)*5
        elif 5000 <= price < 10000:
            price = math.trunc(price/10)*10
        elif 10000 <= price < 50000:
            price = math.trunc(price/50)*50
        elif 50000 <= price < 100000:
            price = math.trunc(price/100)*100
        elif 100000 <= price and type=='KOSDAQ':
            price = math.trunc(price/100)*100
        elif 100000 <= price < 500000 and type=='KOSPI':
            price = math.trunc(price/500)*500
        elif 500000 <= price  and type=='KOSPI':
            price = math.trunc(price/1000)*1000

        return price

    def calcSellPrice(self, high, low, type):
        cl = (high+low)/2
        #price = low+((high-cl)*2*0.8125)
        price = low + ((high - cl) * 2 * 0.9)

        if price<1000:
            price = math.trunc(price)
        elif 1000 <= price < 5000:
            price = math.trunc(price/5)*5
        elif 5000 <= price < 10000:
            price = math.trunc(price/10)*10
        elif 10000 <= price < 50000:
            price = math.trunc(price/50)*50
        elif 50000 <= price < 100000:
            price = math.trunc(price/100)*100
        elif 100000 <= price and type=='KOSDAQ':
            price = math.trunc(price/100)*100
        elif 100000 <= price < 500000 and type=='KOSPI':
            price = math.trunc(price/500)*500
        elif 500000 <= price  and type=='KOSPI':
            price = math.trunc(price/1000)*1000

        return  price

    def calcLossSellPrice(self, high, low, type):
        cl = (high+low)/2
        #price = low+((high-cl)*2*0.8125)
        price = low + ((high - cl) * 2 * 0.25)

        if price<1000:
            price = math.trunc(price)
        elif 1000 <= price < 5000:
            price = math.trunc(price/5)*5
        elif 5000 <= price < 10000:
            price = math.trunc(price/10)*10
        elif 10000 <= price < 50000:
            price = math.trunc(price/50)*50
        elif 50000 <= price < 100000:
            price = math.trunc(price/100)*100
        elif 100000 <= price and type=='KOSDAQ':
            price = math.trunc(price/100)*100
        elif 100000 <= price < 500000 and type=='KOSPI':
            price = math.trunc(price/500)*500
        elif 500000 <= price  and type=='KOSPI':
            price = math.trunc(price/1000)*1000

        return  price

    def calcVIprice(self, open_price, type):
        # price = open_price*1.1
        price = open_price * 1.1
        if price<1000:
            price = math.trunc(price)
        elif 1000 <= price < 5000:
            price = math.trunc(price/5)*5
        elif 5000 <= price < 10000:
            price = math.trunc(price/10)*10
        elif 10000 <= price < 50000:
            price = math.trunc(price/50)*50
        elif 50000 <= price < 100000:
            price = math.trunc(price/100)*100
        elif 100000 <= price and type=='KOSDAQ':
            price = math.trunc(price/100)*100
        elif 100000 <= price < 500000 and type=='KOSPI':
            price = math.trunc(price/500)*500
        elif 500000 <= price  and type=='KOSPI':
            price = math.trunc(price/1000)*1000
        return price

    def setPrice(self, price, type):
        if price<1000:
            price = math.ceil(price)
        elif 1000 <= price < 5000:
            price = (math.ceil(price/5))*5
        elif 5000 <= price < 10000:
            price = (math.ceil(price/10))*10
        elif 10000 <= price < 50000:
            price = (math.ceil(price/50))*50
        elif 50000 <= price < 100000:
            price = (math.ceil(price/100))*100
        elif 100000 <= price and type=='KOSDAQ':
            price = (math.ceil(price/100))*100
        elif 100000 <= price < 500000 and type=='KOSPI':
            price = (math.ceil(price/500))*500
        elif 500000 <= price  and type=='KOSPI':
            price = (math.ceil(price/1000))*1000
        return price

    #def make_Buylist(self, stock_1m, open_price, before_1d_close, type, name):
    def make_Buylist(self, stock_1m, open_price, type, name):

        vi_price = self.calcVIprice(open_price, type)
        #lowPrice = before_1d_close
        lowPrice = 10000000000
        highPrice = 0

        # 실거래정보
        trade_buy_price = 0
        trade_sell_price = 0
        trade_buy_time = 0
        trade_sell_time = 0

        # vi 판별용
        target_VI_price = False
        target_VI_firstcheck = True
        target_VI_price_time = 2

        target_high_price = False
        target_high_firstcheck = True
        target_high_price_time = 1

        target_acc_money = False
        target_money_firstcheck = True
        target_acc_money_time = 0

        relTimecut_num = 0

        for i in range(len(stock_1m)):

            # 최고가와 최저가 갱신시 변수에 넣어준다
            if(lowPrice > stock_1m['저가'][i]):
                lowPrice = stock_1m['저가'][i]
            if(highPrice < stock_1m['고가'][i]):
                highPrice = stock_1m['고가'][i]

            # 구매가 판매가 계산
            buyPrice = self.setPrice(open_price*self.search_price, type)
            sellPrice = self.setPrice(buyPrice*self.target_price, type)
            lossPrice = self.setPrice(buyPrice*self.loss_price, type)
            if(i == 0):
                continue    # 첫봉에선 무조건 구매 안함

            # 시가가 전일종가보다 9% 이하인 종목만 매수대상에 넣음
            # if before_1d_close * st_option.open_gap_price <= open_price:
            #     return pd.DataFrame(columns=self.trade_columns)



            # 거래대금 조건을 첫번째로 만족한 시간 기록
            if (stock_1m['누적거래대금'][i] >= self.tradeAccMoney and target_money_firstcheck==True):
                target_money_firstcheck = False
                target_acc_money = True
                target_acc_money_time = stock_1m['시간'][i]


            # 가격 조건을 첫번째로 만족한 시간 기록
            # 이건 분봉 상승 조건
            # if (stock_1m['종가'][i]/stock_1m['시가'][i] >= st_option.search_price and target_high_firstcheck == True):
            #     target_high_firstcheck = False
            #     target_high_price = True
            #     target_high_price_time = stock_1m['시간'][i]


            # 이건 VI조건   and stock_1m['고가'][i] == stock_1m['종가'][i]
            # if(stock_1m['고가'][i] >= vi_price  and target_VI_firstcheck==True ):
            #     target_VI_firstcheck = False
            #     target_VI_price = True
            #     target_VI_price_time = stock_1m['시간'][i]
            #     continue

            # 돌파조건
            if (stock_1m['고가'][i] >= buyPrice and target_high_firstcheck == True):
                target_high_firstcheck = False
                target_high_price = True
                target_high_price_time = stock_1m['시간'][i]


            # 아직 매수한 상태가 아니라면 조건이 맞는지 확인후 매수
            if trade_buy_price == 0:
                # if (target_acc_money_time<=target_VI_price_time and target_high_price_time <= target_VI_price_time
                #     and target_VI_price == True and target_high_price == True
                #     and target_acc_money == True ):
                if (target_acc_money_time <= target_high_price_time and  target_acc_money == True and target_high_price == True):
                    if stock_1m['고가'][i] >= buyPrice :
                        # 시간 타임컷
                        if stock_1m['시간'][i] >= self.buy_timecut:
                            return pd.DataFrame(columns=self.trade_columns)

                        if buyPrice >= stock_1m['저가'][i]:
                            trade_buy_price = buyPrice
                        else:
                            trade_buy_price = stock_1m['시가'][i]

                        trade_buy_time = stock_1m['시간'][i]
                        relTimecut_num = i+st_option.relative_timecut
                        continue

            # 매수한 상태고 매도가에 도달하거나 타임컷에 도달하면 매도
            # or lossPrice > stock_1m['저가'][i] <- 이거 추가하면 손절거는거임
            if trade_buy_price != 0 and trade_sell_price == 0:
                if(stock_1m['고가'][i]>= sellPrice   or relTimecut_num==i or stock_1m['시간'][i]>=st_option.sell_timecut):     # 타임컷은 종배로함 (len(stock_1m) - 1)
                    if stock_1m['고가'][i]>= sellPrice:
                        trade_sell_price = sellPrice
                    # elif lossPrice > stock_1m['저가'][i]:
                    #     trade_sell_price = lossPrice
                    elif relTimecut_num==i:
                        trade_sell_price = stock_1m['종가'][i]
                    elif stock_1m['시간'][i]>=st_option.sell_timecut :
                        trade_sell_price = stock_1m['종가'][i]
                    trade_sell_time = stock_1m['시간'][i]

                    return pd.DataFrame({'종목코드': stock_1m['종목코드'][i], '종목명': name,
                                             '매수일': stock_1m['날짜'][i], '매수시간': trade_buy_time, '매수가격': trade_buy_price,
                                             '매도일':stock_1m['날짜'][i], '매도시간': trade_sell_time, '매도가격': trade_sell_price,
                                         '수익률':np.nan, '수익금':np.nan},
                                             columns=self.trade_columns,
                                            index = [0])

        # 아무것도 만족하지 못했으면 빈 데이터프레임 리턴
        return pd.DataFrame(columns=self.trade_columns)




# 백테스트 초기화
st_option = Strategy_Option()

# MSSQL 연결
conn_1m = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=STOCK1m_3;UID=아이디;PWD=비번')
cursor_1m = conn_1m.cursor()
conn_1d = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=STOCK1d_2;UID=아이디;PWD=비번')
cursor_1d = conn_1d.cursor()
conn_market = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=MARKET_TYPE;UID=아이디;PWD=비번')
cursor_market = conn_market.cursor()

print('시작', datetime.now())
start_time = datetime.now()

# 시가보다 9.5% 상승하고, 거래대금을 만족했는지 일봉 데이터로 먼저 검증해서 내역 뽑아옴
sql_date = '날짜 >= ' + str(st_option.Start_date) + ' and 날짜 <= ' + str(st_option.End_date) + " "
sql_price = 'and 고가 >= (시가*'+str(st_option.search_price)+')' + " "
sql_volume = "and 거래대금 >= " + str(st_option.tradeAccMoney) + " "
data_1d = pd.read_sql("select * from stock_1d WHERE 종목코드 like 'A%' and " + sql_date + sql_price + sql_volume +
                   "ORDER BY 날짜 ASC, 종목코드 ASC", conn_1d)

# data_1d = pd.read_sql("""
#     select
#         *,
#         (select 종가 from stock_1d d3 where d3.날짜 = big_t.전일날짜 and d3.종목코드 = big_t.종목코드) 전일종가
#     from (
#      select
#         *,
#         (select 날짜 from stock_1d d2 where d2.날짜 < d1.날짜 and d2.종목코드=d1.종목코드 order by 날짜 desc offset 0 rows fetch next 1 rows only) 전일날짜
#         from stock_1d d1 WHERE 종목코드 like 'A%' and """+ sql_date + sql_price + sql_volume +
#     ") big_t where 전일날짜 IS not NULL "
#      +
#     "ORDER BY 날짜 ASC, 종목코드 ASC", conn_1d)

# 종가를 한칸 내려서 새로운 컬럼에 전일종가 기록함
#data_1d['전일종가'] = data_1d['종가'].shift(1)

print(data_1d)

# 거래일에 포함하는 날짜들 추출
dates = data_1d['날짜'].sort_values(ascending=True).unique()



# 자산 날짜 만들기
st_option.Money = pd.Series(0., index=dates)
#st_option.Money[dates[0]] = st_option.Start_Money

st_option.Equity = pd.Series(0., index=dates)
#st_option.Equity[dates[0]] = st_option.Start_Money

current_money = st_option.Start_Money

st_option.set_equity(dates[0], current_money)

# for문 돌면서 거래일에 해당하는 거래들 처리하기 시작
for date, i in zip(dates, range(len(dates))):
    # if(i==0):
    #     continue    # 전일종가 가져오려고 첫날은 그냥 넘김

    print(date,'백테중... ',i,'/',len(dates))

    # 일봉 전체에서 오늘 거래할 종목들만 filtered_1d에 넣음
    filtered_1d = data_1d[data_1d['날짜']==date]
    filtered_1d.reset_index(drop=True, inplace=True)
    #filtered_1d = filtered_1d[['종목코드','날짜','시가','전일종가']]
    filtered_1d = filtered_1d[['종목코드', '날짜', '시가']]

    # 매수조건 만족한 애들 넣어줄 빈 buy_list 선언
    buy_list = pd.DataFrame(columns=st_option.trade_columns)

    # 특정일에 조건 만족하는 모든종목 buylist에 추가
    for j in range(len(filtered_1d)):
        tablename = filtered_1d['종목코드'][j]

        # 1분봉 데이터 가져오기
        stock_1m = pd.read_sql("select * from "+ tablename +" where 날짜=" + str(filtered_1d['날짜'][j]) + " ORDER BY 시간 ASC",conn_1m)
        if(stock_1m.empty):
            continue

        # 종목이 코스피인지 코스닥인지 확인
        # cursor_market.execute("select 타입, 종목명 from market_type where 종목코드='" + filtered_1d['종목코드'][j]+"'")
        # market_type = cursor_market.fetchone()
        market_type = pd.read_sql("select 타입, 종목명 from market_type where 종목코드='" + filtered_1d['종목코드'][j]+"'",conn_market)

        if(market_type.empty): # 이거 우선주종목명이 끝에 K붙게 바뀜 수정필요 지금은 우선주 무조건 패스하게 해둔거임
            print('종목타입 조회 실패:', filtered_1d['종목코드'][j], filtered_1d['날짜'][j])
            continue

        # 오늘 살 리스트 가져오기
        # buy_stock = st_option.make_Buylist(stock_1m, filtered_1d['시가'][j], filtered_1d['전일종가'][j],market_type[0], market_type[1])
        # buy_stock = st_option.make_Buylist(stock_1m, filtered_1d['시가'][j], filtered_1d['전일종가'][j], market_type['타입'][0],
        #                                    market_type['종목명'][0])
        buy_stock = st_option.make_Buylist(stock_1m, filtered_1d['시가'][j], market_type['타입'][0],
                                           market_type['종목명'][0])
        if buy_stock.empty:
            continue

        # buy_list에 살 목록들 추가하기
        buy_list = pd.concat([buy_list, buy_stock], ignore_index=True)


    # 종목당 얼마만큼 금액쓸지 정하기
    buy_money = math.trunc(current_money * st_option.buy_ratio)

    # 매수내역 리스트 옵션 클래스에 추가해두기
    buy_list = buy_list[buy_list['매수가격']<=buy_money]    # 1종목당 예산보다 매수가격이 비싼애들 제거
    buy_list = buy_list.sort_values(by='매수일' and '매수시간').head(st_option.maximum_buy)
    buy_list = buy_list.reset_index(drop=True)

    # 사는거, 파는거 처리
    for k in buy_list.index:
        buy_price = buy_list['매수가격'][k]
        sell_price = buy_list['매도가격'][k]
        amount = math.trunc(buy_money/buy_price * (1/(1+st_option.Fee)))


        # 0으로 나누는 오류 있어서 테스트용으로 만듬
        if buy_price ==0 or sell_price ==0 or amount == 0:
            print(buy_list['종목코드'][k],buy_list['종목명'][k],buy_list['매수일'][k],buy_list['매수시간'][k],buy_price,sell_price,amount,buy_money)

        # 사는거 처리
        current_money -= buy_price * amount
        # 파는거 처리
        current_money += math.trunc((sell_price * amount) * (1-(st_option.Fee + st_option.Tax)))

        buy_list.loc[k,'수익률'] = round(((math.trunc((sell_price * amount) * (1-(st_option.Fee + st_option.Tax))) / (buy_price * amount))-1) * 100, 2)
        buy_list.loc[k,'수익금'] = math.trunc((sell_price * amount) * (1-(st_option.Fee + st_option.Tax))) - (buy_price * amount)

    st_option.set_equity(date, current_money)

    st_option.trade_list = pd.concat([st_option.trade_list, buy_list], ignore_index=True)

    # if(i==2):
    #     break
print(st_option.trade_list)

print('자산:', st_option.Equity)
print('돈:', st_option.Money)
print('CAGR: ', st_option.get_cagr()*100, '    MDD: ', st_option.get_mdd()*100)



print('끝', datetime.now())
end_time = datetime.now()
print(end_time - start_time)

st_option.trade_list.to_excel("돌파 day_trading"+end_time.strftime("%Y%m%d_%H_%M_%S")+".xlsx")
#ax.plot(dates, st_option.Equity)
#plot_equity()

date_line = []
for date in dates:
    date = datetime.strptime(str(date),"%Y%m%d")
    date_line.append(date.strftime((str(date.strftime("%Y%m%d")))))

plt.title('breakthrough')
plt.plot(date_line, st_option.Equity)
plt.axis(option='auto')
plt.xlabel('time')
fig.autofmt_xdate()
plt.show()


