#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import datetime
import pyodbc
import pandas as pd
import subprocess
import requests
import calendar
from ftplib import FTP_TLS
from datetime import date,timedelta

# 25/02/27 v1.42 微修正
version = "1.42"

# TODO: 順位関数の共通化

appdir = os.path.dirname(os.path.abspath(__file__))

conn_temp = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=xxxxxx;'
    )
df = ""
out = ""
templatefile = appdir + "./bktemplate.htm"
resultfile = appdir + "./bookdr.htm"
conffile = appdir + "./bookdr.conf"
dbfile = ""
start_year = 1990
end_year = 0        # 集計する最終年
accdata = {}        # 現在月までの累積データ   キー  年  値  リスト
lastdate = ""
star_icon = '<i class="fa-solid fa-star" style="color: #73e65a;"></i>'
book_icon = '<i class="fa-solid fa-book" style="color: #73e65a;"></i>'
info_icon = '<i class="fa-solid fa-circle-info" style="color: #73e65a;"></i>'
yen_icon = '<i class="fa-solid fa-sack-dollar" style="color: #73e65a;"></i>'
pixela_url = ""
pixela_token = ""

def main_proc() :
    global end_year
    date_settings()
    end_year = today_yy  # 集計する最終年

    read_config()
    read_database()
    accumulate()
    calc_rank_month()
    create_df_year()
    parse_template()
    if debug == 1 :
        return
    _  = subprocess.run((browser, resultfile))
    ftp_upload()
    post_pixela()

def read_config() :
    global ftp_host,ftp_user,ftp_pass,ftp_url,dbfile,browser,pixela_url,pixela_token,debug
    if not os.path.isfile(conffile) :
        return
    conf = open(conffile,'r', encoding='utf-8')
    dbfile = conf.readline().strip()
    browser = conf.readline().strip()
    ftp_host = conf.readline().strip()
    ftp_user = conf.readline().strip()
    ftp_pass = conf.readline().strip()
    ftp_url = conf.readline().strip()
    pixela_url = conf.readline().strip()
    pixela_token = conf.readline().strip()
    debug  = int(conf.readline().strip())
    conf.close()

def read_database():
    global df,lastdate 
    conn_str = conn_temp.replace("xxxxxx",dbfile)
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    sql = 'SELECT * FROM MAIN'

    rows = cursor.execute(sql).fetchall()
    date_list = []
    price_list = []
    page_list = []
    lib_list = []
    title_list = []
    for r in rows :
        str = f'{r[1]}/{r[2]:02}/{r[3]:02}'
        dte = datetime.datetime.strptime(str, '%Y/%m/%d')
        title = r[4]
        lib = r[8]
        page = r[10]
        price = r[11]
        libflg = 0 
        if lib == 'L' :
            libflg = 1
        date_list.append(dte)
        price_list.append(price)
        page_list.append(page)
        lib_list.append(libflg)
        title_list.append(title)
        lastdate = str

    df = pd.DataFrame(list(zip(date_list,price_list,page_list,lib_list,title_list))
        , columns = ['date','price','page','lib','title'])

#   価格ランキング
def rank_price():
    df_s = df.sort_values(by=['price'],ascending=False)
    rank_price_output(df_s,20)

#   価格ランキング 365日
def rank_price_year():
    #  1年前の同月(を含まない)以降からのランキング
    target_mm = today_mm + 1 
    if target_mm == 13 :
        target_mm = 1 
    target_yy = end_year -1 
    target_df = df[df['date'] >= datetime.datetime(target_yy,target_mm,1)]
    df_s = target_df.sort_values(by=['price'],ascending=False)
    rank_price_output(df_s,20)

#   価格ランキングの表示   上位 n 個を表示する
def rank_price_output(target_df,n) :
    i = 0 
    for _, row in target_df.iterrows():
        i = i+1
        libstr = ""
        if row.lib == 1 :
            libstr = star_icon
        dd = row.date.strftime("%y/%m/%d")
        out.write(f'<tr><td align="right">{i}</td><td>{dd}</td><td>{row.title}</td>'
                  f'<td align="right">{row.price:5.0f}</td><td align="center">{libstr}</td></tr>')
        if i == n : 
            return

#   ページランキング
def rank_page():
    df_s = df.sort_values(by=['page'],ascending=False)
    rank_page_output(df_s,20)

#   ページランキング 365日
def rank_page_year():
    #  1年前の同月(を含まない)以降からのランキング
    target_mm = today_mm + 1 
    if target_mm == 13 :
        target_mm = 1 
    target_yy = end_year -1 
    target_df = df[df['date'] >= datetime.datetime(target_yy,target_mm,1)]
    df_s = target_df.sort_values(by=['page'],ascending=False)
    rank_page_output(df_s,20)

#  月ごとの  ページ、価格 のデータフレーム df_month を作成する
def calc_rank_month() :
    global df_month
    date_list = []
    count_list = []      #  月冊数
    page_list = []
    price_list = []
    ave_count_list = []   #  30日平均冊数
    ave_page_list = []   #  日平均ページ数
    ave_price_list = []   #  日平均ページ数
    for yy in range(1994,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        for mm in range(1,13) : 
            if yy == end_year and mm > today_mm :
                break
            dfmm = dfyy[dfyy['date'].dt.month == mm]
            count = len(dfmm)
            count_list.append(count)
            pg = dfmm['page'].sum()
            page_list.append(pg)
            pr = dfmm['price'].sum()
            price_list.append(pr)
            date_list.append(yy*100+mm)
            ave = pg / calendar.monthrange(yy, mm)[1]
            ave_price = pr
            ave_count = count
            if yy == end_year and mm == today_mm  :   #  今月
                ave = pg / today_dd
                ave_price = pr / today_dd * 30
                ave_count = count / today_dd * 30

            ave_page_list.append(ave)
            ave_price_list.append(ave_price)
            ave_count_list.append(ave_count)

    df_month = pd.DataFrame(list(zip(date_list,count_list,page_list,price_list,ave_count_list,ave_page_list,ave_price_list))
        , columns = ['date','count','page','price','ave_count','ave_page','ave_price'])

#  年ごとの  ページ、価格 のデータフレーム df_year を作成する
def create_df_year() :
    global df_year
    today = date.today()
    start_date = date(today.year, 1, 1)
    days_year = (today - start_date).days   #  1/1 からの日数

    date_list = []
    count_list = []      #  年冊数
    ave_count_list = []      #  年冊数
    page_list = []
    ave_page_list = []   #  日平均ページ数
    price_list = []
    ave_price_list = []
    for yy in range(1994,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        count = len(dfyy)
        count_list.append(count)
        pg = dfyy['page'].sum()
        page_list.append(pg)
        pr = dfyy['price'].sum()
        price_list.append(pr)
        ave_count = count / 12
        ave = pg / 365
        ave_pr = pr / 12
        if yy == end_year :
            ave_count = count / days_year * 30
            ave = pg / days_year
            ave_pr = pr / days_year * 30
        ave_page_list.append(ave)
        ave_price_list.append(ave_pr)
        ave_count_list.append(ave_count)
        date_list.append(yy)
            
    df_year = pd.DataFrame(list(zip(date_list,count_list,page_list,price_list,ave_count_list,ave_page_list,ave_price_list))
        , columns = ['date','count','page','price','ave_count','ave_page','ave_price'])

# 今年の冊数順位
def  cur_year_count_rank() :
    order = int(df_year['count'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_year)
    return order,count

# 今年の平均冊数順位
def  cur_year_ave_count_rank() :
    order = int(df_year['ave_count'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_year)
    return order,count

# 今年のページ順位
def  cur_year_page_rank() :
    order = int(df_year['page'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_year)
    page = df_year['page'].iloc[-1]
    return order,count,page

# 今年の平均ページ順位
def  cur_year_ave_page_rank() :
    order = int(df_year['ave_page'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_year)
    ave_page = df_year['ave_page'].iloc[-1]
    return order,count,ave_page

# 今年の価格順位
def  cur_year_price_rank() :
    order = int(df_year['price'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_year)
    price = int(df_year['price'].iloc[-1])
    return order,count,price

# 今年の月平均価格順位
def  cur_year_ave_price_rank() :
    order = int(df_year['ave_price'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_year)
    price = int(df_year['ave_price'].iloc[-1])
    return order,count,price

# 今月の冊数順位
def  cur_month_count_rank() :
    order = int(df_month['count'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_month)
    return order,count

# 今月の平均冊数順位
def  cur_month_ave_count_rank() :
    order = int(df_month['ave_count'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_month)
    return order,count

# 今月のページ順位
def  cur_month_page_rank() :
    order = int(df_month['page'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_month)
    page = int(df_month['page'].iloc[-1])
    return order,count,page

# 今月の平均ページ順位
def  cur_month_ave_page_rank() :
    order = int(df_month['ave_page'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_month)
    page = int(df_month['ave_page'].iloc[-1])
    return order,count,page

# 今月の価格順位
def  cur_month_price_rank() :
    order = int(df_month['price'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_month)
    price = int(df_month['price'].iloc[-1])
    return order,count,price

# 今月の平均価格順位
def  cur_month_ave_price_rank() :
    order = int(df_month['ave_price'].rank(method='min',ascending=False).iloc[-1])  # 最終行(=今月)のindexを取得
    count = len(df_month)
    price = int(df_month['ave_price'].iloc[-1])
    return order,count,price

#  月別ページランキングの表示
def rank_page_month(flg) :
    df_page_month_sort = df_month.sort_values(by=['page'],ascending=False)
    rank_month_com(flg,df_page_month_sort,0)

#  月別価格ランキング
def rank_price_month(flg) :
    # flg 1 の時 1 .. 10 位を表示、 2 の時 11 .. 20 位を表示  3  21 - 31 位を表示
    df_price_month_sort = df_month.sort_values(by=['price'],ascending=False)
    rank_month_com(flg,df_price_month_sort,1)

def rank_month_com(flg,df,kind) :
    # kind  0 ... page   1 .. price

    i = 0
    for _, row in df.iterrows():
        i = i+1
        if flg == 1 :
            if i > 10 :
                break
        elif flg == 2 :
            if i <= 10 :
                continue
            if i > 20 :
                break
        else :
            if i <= 20 :
                continue

        yymm = int(row.date) 
        yy = int(yymm / 100 )
        mm = yymm % 100
        str_yymm = f'{yy}/{mm:02}'
        if yymm  == today_yymm :     #  今月は赤字
            str_yymm = f'<span class=red>{str_yymm}</span>'
        if kind == 0 :
            out.write(f'<tr><td align="right">{i}</td><td>{str_yymm}</td><td align="right">{int(row.page):,}</td></tr>')
        else :
            out.write(f'<tr><td align="right">{i}</td><td>{str_yymm}</td><td align="right">{int(row.price):,}</td></tr>')
        if i == 30 : 
            break

def rank_page_output(target_df,n) :
    i = 0
    for _, row in target_df.iterrows():
        i = i+1
        libstr = ""
        if row.lib == 1 :
            libstr = star_icon
        dd = row.date.strftime("%y/%m/%d")
        out.write(f'<tr><td align="right">{i}</td><td>{dd}</td><td>{row.title}</td>'
                  f'<td align="right">{row.page:5.0f}</td><td align="center">{libstr}</td></tr>')
        if i == n : 
            return

#   年別の現在月での累積データ
def accumulate() :
    global accdata
    for yy in range(start_year+4,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        
        dfmmacc = ""
        acclist = []
        for mm in range(1,today_mm+1) :
            dfmm = dfyy[dfyy['date'].dt.month == mm]
            if mm == 1 :
                dfmmacc = dfmm
            else :
                dfmmacc = pd.concat([dfmmacc, dfmm])
        n = len(dfmmacc)    # 冊数
        lib = dfmmacc['lib'].sum()     # 図書館冊数
        acclist.append(n)
        acclist.append(lib)
        page_sum = dfmmacc['page'].sum()
        page_mean = dfmmacc['page'].mean()
        price_sum = dfmmacc['price'].sum()
        price_mean = dfmmacc['price'].mean()
        acclist.append(page_sum)
        acclist.append(page_mean)
        acclist.append(price_sum)
        acclist.append(price_mean)
        accdata[yy] = acclist

def acc_table():
    for yy in range(start_year+4,end_year+1) :
        acclist = accdata[yy]
        n = acclist[0]
        lib = acclist[1]
        
        out.write(f"<tr><td>{yy}</td><td align='right'>{n}</td>"
                  f"<td align='right'>{acclist[2]:5.0f}</td>"
                  f"<td align='right'>{acclist[3]:5.1f}</td>"
                  f"<td align='right'>{acclist[4]:5.0f}</td>"
                  f"<td align='right'>{acclist[5]:5.1f}</td>"
                  f"<td align='right'>{lib:5.0f}</td><td align='right'>{lib/n*100:3.1f}</td>"
                  f"</tr>\n")

def acc_graph(): 
    for yy in range(start_year+4,end_year+1) :
        acclist = accdata[yy]
        n = acclist[0]
        out.write(f"['{yy:02}',{n}],") 

#  月別データ
def month_table() :
    for yy in range(end_year-2,end_year+1) :    # 3年分
        dfyy = df[df['date'].dt.year == yy]
        for mm in range(1,13) :                 #  1 - 12 月
            if yy == end_year and mm > today_mm :
                break
            dfmm = dfyy[dfyy['date'].dt.month == mm]
            n = len(dfmm)               # 冊数
            lib = dfmm['lib'].sum()     # 図書館冊数
            librate = 0 
            if n != 0 :
                librate = lib/n*100
            
            if n == 0 :
                page_mean = 0 
                price_mean = 0 
            else :
                page_mean = dfmm['page'].mean()
                price_mean = dfmm['price'].mean()

            out.write(f"<tr><td>{yy}/{mm:02}</td><td align='right'>{n}</td>"
                    f"<td align='right'>{dfmm['page'].sum():5.0f}</td>"
                    f"<td align='right'>{page_mean:5.1f}</td>"
                    f"<td align='right'>{dfmm['price'].sum():5.0f}</td>"
                    f"<td align='right'>{price_mean:5.1f}</td>"
                    f"<td align='right'>{lib:5.0f}</td><td align='right'>{librate:3.1f}</td>"
                    f"</tr>\n")

def year_table() :
    global price_year_ave,librate_year_ave
    for yy in range(start_year,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        n = len(dfyy)               # 冊数
        lib = dfyy['lib'].sum()     # 図書館冊数
        librate = 0 
        if n != 0 :
            librate = lib/n*100

        out.write(f"<tr><td>{yy}</td><td align='right'>{n}</td>"
                  f"<td align='right'>{dfyy['page'].sum():5,.0f}</td>"
                  f"<td align='right'>{dfyy['page'].mean():5.1f}</td>"
                  f"<td align='right'>{dfyy['price'].sum():5,.0f}</td>"
                  f"<td align='right'>{dfyy['price'].mean():5.1f}</td>"
                  f"<td align='right'>{lib:5.0f}</td><td align='right'>{librate:3.1f}</td>"
                  f"</tr>\n")

def monthly_graph():
    for yy in range(end_year-2,end_year+1) :    # 3年分
        dfyy = df[df['date'].dt.year == yy]
        for mm in range(1,13) :
            dfmm = dfyy[dfyy['date'].dt.month == mm]
            yy2 = yy - 2000
            out.write(f"['{yy2:02}{mm:02}',{len(dfmm)}],") 

def year_graph():
    for yy in range(start_year,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        yy2 = yy % 100
        out.write(f"['{yy2:02}',{len(dfyy)}],") 

def year_price_graph():
    price_year_ave = {}
    for yy in range(1994,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        price_year_ave[yy] = dfyy['price'].mean()
        yy2 = yy % 100
        out.write(f"['{yy2:02}',{price_year_ave[yy]}],") 

def year_librate_graph():
    librate_year_ave = {}
    for yy in range(1994,end_year+1) :
        dfyy = df[df['date'].dt.year == yy]
        n = len(dfyy)               # 冊数
        lib = dfyy['lib'].sum()     # 図書館冊数
        librate_year_ave[yy] = lib/n*100
        yy2 = yy % 100
        out.write(f"['{yy2:02}',{librate_year_ave[yy]}],") 

def date_settings():
    global  today_date,today_mm,today_dd,today_yy,lastdate,today_datetime,today_yymm
    today_datetime = datetime.datetime.today()
    today_date = datetime.date.today()
    today_mm = today_date.month
    today_dd = today_date.day
    today_yy = today_date.year
    today_yymm = today_yy * 100 + today_mm  # yyyymm の形式にする
    #print(today_yymm)
    #lastdate = today_date - timedelta(days=1)

def today(s):
    d = today_datetime.strftime("%m/%d %H:%M")
    s = s.replace("%today%",d)
    out.write(s)

def summary() :
    num_all = len(df)              # 全冊数
    page_all = df['page'].sum()
    price_all = df['price'].sum()

    dfyy = df[df['date'].dt.year == end_year]
    num_year  = len(dfyy)          # 今年の冊数
    page_year = dfyy['page'].sum()
    price_year = dfyy['price'].sum()

    dfmm = dfyy[dfyy['date'].dt.month == today_mm]
    num_month  = len(dfmm)         # 今月の冊数
    page_month = dfmm['page'].sum()
    price_month = dfmm['price'].sum()

    start_date = date(today_yy, 1, 1)
    days_year = (today_date - start_date).days
    start_date = date(today_yy, today_mm, 1)
    days_month = (today_date - start_date).days + 1
    start_date = date(1990, 4, 1)
    days_all = (today_date - start_date).days

    count_order,count = cur_month_count_rank()
    ave_page_order,count,ave_page = cur_month_ave_page_rank()
    page_order,count,cur_page = cur_month_page_rank()
    price_order,count,cur_price = cur_month_price_rank()
    ave_price_order,count,cur_price = cur_month_ave_price_rank()
    ave_count_order,count = cur_month_ave_count_rank()
        
    out.write('<tr>')
    out.write(f'<td>今月</td><td align="right">{num_month}</td><td align="right">{count_order}/{count}</td>')
    out.write(f'<td align="right">{num_month/days_month*30:.2f}</td>')
    out.write(f'<td align="right">{ave_count_order}/{count}</td>')
    out.write(f'<td align="right">{page_month:,.0f}</td><td align="right">{page_order}/{count}</td>')
    out.write(f'<td align="right">{page_month/days_month:.0f}</td><td align="right">{ave_page_order}/{count}</td>')
    out.write(f'<td align="right">{price_month:,.0f}</td><td align="right">{price_order}/{count}</td>')
    out.write(f'<td align="right">{price_month/days_month*30:,.0f}</td>')
    out.write(f'<td align="right">{ave_price_order}/{count}</td>')
    out.write('</tr>\n')

    count_order,count = cur_year_count_rank()
    ave_count_order,count = cur_year_ave_count_rank()
    ave_page_order,count,ave_page = cur_year_ave_page_rank()
    page_order,count,page = cur_year_page_rank()
    price_order,count,acc_price = cur_year_price_rank()
    ave_price_order,count,acc_price = cur_year_ave_price_rank()

    out.write('<tr>')
    out.write(f'<td>今年</td><td  align="right">{num_year}</td><td align="right">{count_order}/{count}</td>')
    out.write(f'<td align="right">{num_year/days_year*30:.2f}</td>')
    out.write(f'<td align="right">{ave_count_order}/{count}</td>')
    out.write(f'<td align="right">{page_year:,.0f}</td><td  align="right">{page_order}/{count}</td>')
    out.write(f'<td align="right">{page_year/days_year:.0f}</td><td  align="right">{ave_page_order}/{count}</td>')
    out.write(f'<td align="right">{price_year:,.0f}</td><td  align="right">{price_order}/{count}</td>')
    out.write(f'<td align="right">{price_year/days_year*30:,.0f}</td>')
    out.write(f'<td align="right">{ave_price_order}/{count}</td>')
    out.write('</tr>\n')

    out.write('<tr>')
    out.write(f'<td>総合</td><td align="right">{num_all}</td><td></td><td  align="right">{num_all/days_all*30:.2f}</td>')
    out.write(f'<td></td>')
    out.write(f'<td  align="right">{page_all:,.0f}</td><td></td><td  align="right">{page_all/days_all:.0f}</td></td><td>')
    out.write(f'<td  align="right">{price_all:,.0f}</td><td></td><td  align="right">{price_all/days_all*30:,.0f}</td>')
    out.write(f'<td align="right"></td>')
    out.write('</tr>\n')

def post_pixela() :
    post_days = 14      #  最近の何日をpostするか
    headers = {}
    headers['X-USER-TOKEN'] = pixela_token
    dte = datetime.datetime.strptime(lastdate, '%Y/%m/%d')
    startdate =  dte - timedelta(post_days)
    for _, row in df.iterrows():
        chk_date = row.date
        if chk_date < startdate :
            continue
        data = {}
        dd = chk_date.strftime('%Y%m%d') 
        data['date'] = dd
        data['quantity'] = "1"
        response = requests.post(url=pixela_url, json=data, headers=headers,verify=False)

def parse_template() :
    global out 
    f = open(templatefile , 'r', encoding='utf-8')
    out = open(resultfile,'w' ,  encoding='utf-8')
    for line in f :
        # if "%lastdate%" in line :
        #     curdate(line)
        #     continue
        if "%month_table%" in line :
            month_table()
            continue
        if "%monthly_graph%" in line :
            monthly_graph()
            continue
        if "%year_table%" in line :
            year_table()
            continue
        if "%year_graph%" in line :
            year_graph()
            continue
        if "%accumulate%" in line :
            acc_table()
            continue
        if "%acc_graph%" in line :
            acc_graph()
            continue
        if "%rank_price%" in line :
            rank_price()
            continue
        if "%rank_price_year%" in line :
            rank_price_year()
            continue
        if "%rank_page%" in line :
            rank_page()
            continue
        if "%rank_page_year%" in line :
            rank_page_year()
            continue
        if "%rank_page_month1%" in line :
            rank_page_month(1)
            continue
        if "%rank_page_month2%" in line :
            rank_page_month(2)
            continue
        if "%rank_page_month3%" in line :
            rank_page_month(3)
            continue
        if "%rank_price_month1%" in line :
            rank_price_month(1)
            continue
        if "%rank_price_month2%" in line :
            rank_price_month(2)
            continue
        if "%rank_price_month3%" in line :
            rank_price_month(3)
            continue
        if "%cur_month%" in line :
            out.write(f'{today_mm} 月現在')
            continue
        if "%today%" in line :
            today(line)
            continue
        if "%year_price_graph%" in line :
            year_price_graph()
            continue
        if "%year_librate_graph%" in line :
            year_librate_graph()
            continue

        if "%version%" in line :
            s = line.replace("%version%",version)
            out.write(s)
            continue
        if "%lastdate%" in line :
            s = line.replace("%lastdate%",lastdate)
            out.write(s)
            continue
        if "%summary%" in line :
            summary()
            continue

        out.write(line)

    f.close()
    out.close()

def ftp_upload() : 
    with FTP_TLS(host=ftp_host, user=ftp_user, passwd=ftp_pass) as ftp:
        ftp.storbinary('STOR {}'.format(ftp_url), open(resultfile, 'rb'))

def curdate():
    pass

#-----------------------------------
main_proc()
