#%%
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import numpy as np
import math
import pandas as pd
import pandas.tseries.offsets as offsets
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()

# グラフ文字化け対策
mpl.rcParams['font.family'] = 'Noto Sans CJK JP'
plt.rcParams['grid.linestyle']='--'
plt.rcParams['grid.linewidth'] = 0.5

# グラフのフォントサイズ
plt.rcParams["font.size"] = 12

def readOndotoriFile(filename):
    '''
    おんどとりのデータを読み込む関数
    '''
    df = pd.read_excel(filename, usecols=[0, 2, 3], skiprows=[1,2], converters={'No.1':np.float64, 'No.2':np.float64,})
    df = df.rename(columns={'No.1': '温度', 'No.2': '湿度'})
    df['Date/Time'] = pd.to_datetime(df['Date/Time'])
    df.set_index('Date/Time', inplace=True)

    # 確認用
    # print(df)
    # print(df["温度"])
    # print(df.loc['2020-07-28 10:15:00', '温度'])

    return df


def readHoboFile(filename):
    '''
    Onset Hoboのデータを読み込む関数
    '''
    df = pd.read_excel(filename, usecols=[1, 2, 3, 4, 5], skiprows=[0,1], converters={'床':np.float64, '天井':np.float64,'壁':np.float64,'柱':np.float64})
    df['Unnamed: 1'] = df['Unnamed: 1'].str.replace("午前", "AM")
    df['Unnamed: 1'] = df['Unnamed: 1'].str.replace("午後", "PM")
    df['Date/Time'] = pd.to_datetime(df['Unnamed: 1'])
    df.set_index('Date/Time', inplace=True)
    df = df.drop(columns='Unnamed: 1')

    return df


def makeFigure_Ondotori(area_name,L150,L1500,L2200,B1500,K1500,date_span_start,data_span_end):

    fig = plt.figure(figsize=(10,4.5))
    fig.subplots_adjust(left=0.06, bottom=0.09, right=0.98, top=0.92, wspace=0.20, hspace=0.20)
    fig.suptitle(area_name + ' 室内温度', fontsize=14)
    ax1 = fig.add_subplot(1,1,1)
    ax1.plot(L150[date_span_start:data_span_end]["温度"],'darkturquoise',label = "主室 150mm")
    ax1.plot(L1500[date_span_start:data_span_end]["温度"],'b',label = "主室 1500mm")
    ax1.plot(L2200[date_span_start:data_span_end]["温度"],'slateblue',label = "主室 2200mm")
    ax1.plot(B1500[date_span_start:data_span_end]["温度"],'tomato',label = "寝室 1500mm")
    ax1.plot(K1500[date_span_start:data_span_end]["温度"],'g',label = "台所 1500mm")
    ax1.set_ylabel("室温[℃]", fontsize=14)
    ax1.xaxis.set_major_formatter(DateFormatter('%m/%d %H時'))
    ax1.set_ylim([24,36])
    ax1.legend()
    ax1.grid()
    fig.savefig(area_name + "_01室内温度.png")

    fig = plt.figure(figsize=(10,4.5))
    fig.subplots_adjust(left=0.06, bottom=0.09, right=0.98, top=0.92, wspace=0.20, hspace=0.20)
    fig.suptitle(area_name + ' 室内湿度', fontsize=14)
    ax1 = fig.add_subplot(1,1,1)
    ax1.plot(L150[date_span_start:data_span_end]["湿度"],'darkturquoise',label = "主室 150mm")
    ax1.plot(L1500[date_span_start:data_span_end]["湿度"],'b',label = "主室 1500mm")
    ax1.plot(L2200[date_span_start:data_span_end]["湿度"],'slateblue',label = "主室 2200mm")
    ax1.plot(B1500[date_span_start:data_span_end]["湿度"],'tomato',label = "寝室 1500mm")
    ax1.plot(K1500[date_span_start:data_span_end]["湿度"],'g',label = "台所 1500mm")
    ax1.set_ylabel("湿度[%]", fontsize=14)
    ax1.xaxis.set_major_formatter(DateFormatter('%m/%d %H時'))
    ax1.set_ylim([50,80])
    ax1.legend()
    ax1.grid()
    fig.savefig(area_name + "_02室内温度.png")

def makeFigure_Hobo(area_name, WallTemp, date_span_start, data_span_end):

    fig = plt.figure(figsize=(10,4.5))
    fig.subplots_adjust(left=0.06, bottom=0.09, right=0.98, top=0.92, wspace=0.20, hspace=0.20)
    fig.suptitle(area_name + ' 壁面温度', fontsize=14)
    ax1 = fig.add_subplot(1,1,1)
    ax1.plot(WallTemp[date_span_start:data_span_end]["床"],'darkturquoise',label = "床")
    ax1.plot(WallTemp[date_span_start:data_span_end]["天井"],'b',label = "天井")
    ax1.plot(WallTemp[date_span_start:data_span_end]["壁"],'slateblue',label = "壁")
    ax1.plot(WallTemp[date_span_start:data_span_end]["柱"],'tomato',label = "柱")
    ax1.set_ylabel("表面温度[℃]", fontsize=14)
    ax1.xaxis.set_major_formatter(DateFormatter('%m/%d %H時'))
    ax1.set_ylim([22,36])
    ax1.legend()
    ax1.grid()
    fig.savefig(area_name + "_03壁面温度.png")


#%%

# グラフ化
date_span_start = "2020-07-30 17:00:00"
data_span_end   = "2020-07-31 11:00:00"

# 柳井原2-1のデータ読み込み
Y21_Living_0150  = readOndotoriFile("./data/柳井原団地2-1/1　柳井原2-1　主室　150mm.xlsx")
Y21_Living_1500  = readOndotoriFile("./data/柳井原団地2-1/2　柳井原2-1　主室　1500mm.xlsx")
Y21_Living_2200  = readOndotoriFile("./data/柳井原団地2-1/3　柳井原2-1　主室　2200mm.xlsx")
Y21_Bedroom_1500 = readOndotoriFile("./data/柳井原団地2-1/4　柳井原2-1　寝室　1500mm.xlsx")
Y21_Kitchen_1500 = readOndotoriFile("./data/柳井原団地2-1/5　柳井原2-1　台所　1500mm.xlsx")
Y21_WallTemp     = readHoboFile("./data/柳井原団地2-1/柳井原2-1　kenken_T01.xlsx")

makeFigure_Ondotori("柳井原団地 2-1",Y21_Living_0150,Y21_Living_1500,Y21_Living_2200,Y21_Bedroom_1500,Y21_Kitchen_1500,date_span_start,data_span_end)
makeFigure_Hobo("柳井原団地 2-1", Y21_WallTemp, date_span_start, data_span_end)

# 柳井原2-8のデータ読み込み
Y28_Living_0150  = readOndotoriFile("./data/柳井原団地2-8/6　柳井原2-8　主室　150mm.xlsx")
Y28_Living_1500  = readOndotoriFile("./data/柳井原団地2-8/7　柳井原2-8　主室　1500mm.xlsx")
Y28_Living_2200  = readOndotoriFile("./data/柳井原団地2-8/8　柳井原2-8　主室　2200mm.xlsx")
Y28_Bedroom_1500 = readOndotoriFile("./data/柳井原団地2-8/9　柳井原2-8　寝室　1500mm.xlsx")
Y28_Kitchen_1500 = readOndotoriFile("./data/柳井原団地2-8/10　柳井原2-8　台所　1500mm.xlsx")
Y28_WallTemp     = readHoboFile("./data/柳井原団地2-8/柳井原2-8　kenken_T02.xlsx")

makeFigure_Ondotori("柳井原団地 2-8",Y28_Living_0150,Y28_Living_1500,Y28_Living_2200,Y28_Bedroom_1500,Y28_Kitchen_1500,date_span_start,data_span_end)
makeFigure_Hobo("柳井原団地 2-8", Y28_WallTemp, date_span_start, data_span_end)


# 二万3-4のデータ読み込み
N34_Living_0150  = readOndotoriFile("./data/二万団地3-4/11　二万3-4　主室　150mm.xlsx")
N34_Living_1500  = readOndotoriFile("./data/二万団地3-4/12　二万3-4　主室　1500mm.xlsx")
N34_Living_2200  = readOndotoriFile("./data/二万団地3-4/13　二万3-4　主室　2200mm.xlsx")
N34_Bedroom_1500 = readOndotoriFile("./data/二万団地3-4/14　二万3-4　寝室　1500mm.xlsx")
N34_Kitchen_1500 = readOndotoriFile("./data/二万団地3-4/15　二万3-4　台所　1500mm.xlsx")
N34_WallTemp     = readHoboFile("./data/二万団地3-4/二万団地3-4　kenken_T03.xlsx")

makeFigure_Ondotori("二万団地 3-4",N34_Living_0150,N34_Living_1500,N34_Living_2200,N34_Bedroom_1500,N34_Kitchen_1500,date_span_start,data_span_end)
makeFigure_Hobo("二万団地 3-4", N34_WallTemp, date_span_start, data_span_end)


# 二万4-2のデータ読み込み
N42_Living_0150  = readOndotoriFile("./data/二万団地4-2/16　二万4-2　主室　150mm.xlsx")
N42_Living_1500  = readOndotoriFile("./data/二万団地4-2/17　二万4-2　主室　1500mm.xlsx")
N42_Living_2200  = readOndotoriFile("./data/二万団地4-2/18　二万4-2　主室　2200mm.xlsx")
N42_Bedroom_1500 = readOndotoriFile("./data/二万団地4-2/18　二万4-2　主室　2200mm.xlsx")    ## ここは後で置き換える！！
N42_Kitchen_1500 = readOndotoriFile("./data/二万団地4-2/20　二万4-2　台所　1500mm.xlsx")
N42_WallTemp     = readHoboFile("./data/二万団地4-2/二万団地4-2　kenken_T05.xlsx")

makeFigure_Ondotori("二万団地 4-2",N42_Living_0150,N42_Living_1500,N42_Living_2200,N42_Bedroom_1500,N42_Kitchen_1500,date_span_start,data_span_end)
makeFigure_Hobo("二万団地 4-2", N42_WallTemp, date_span_start, data_span_end)

