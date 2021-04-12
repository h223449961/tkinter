#引入函式庫
from pandas_datareader import data as web
import datetime as dt
import yfinance as yahoo
import matplotlib.pyplot as plt
import tkinter as tk

#導入 yahoo 股資料至本機
yahoo.pdr_override()

#創建視窗
root = tk.Tk()

#獲取當前電腦螢幕的解析度
screenWidth = root.winfo_screenwidth()
screenHeight = root.winfo_screenheight()

#初始運行視窗的座標設置在左上角顯示
x, y = int(screenWidth / 4), int(screenHeight / 4)

#初始化視窗是電腦螢幕的二分之一
width = int(screenWidth / 2)
height = int(screenHeight / 2)

#視窗的大小、初始視窗在螢幕的位置
root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
root.title('xq')

#請輸入股票四碼加後綴地區碼
stock_frame = tk.Frame(root)
stock_frame.pack(side=tk.TOP)
stock_label = tk.Label(stock_frame, text='股票四碼加後綴地區碼')
stock_label.pack(side=tk.LEFT)
stock_entry = tk.Entry(stock_frame)
stock_entry.pack(side=tk.LEFT)

#股票四碼加後綴地區碼範例
exid_frame = tk.Frame(root)
exid_frame.pack(side=tk.TOP)
exid_label = tk.Label(exid_frame, text='範例：2409.tw')
exid_label.pack(side=tk.LEFT)

#起
startTime_frame = tk.Frame(root)
startTime_frame.pack(side=tk.TOP)
startTime_label = tk.Label(startTime_frame, text='起：')
startTime_label.pack(side=tk.LEFT)
input_startdate_var = tk.StringVar()
startdate_widget = tk.Entry(startTime_frame,textvariable = input_startdate_var)
input_startdate_get = input_startdate_var.set(input_startdate_var.get())
startdate_widget.pack(side = tk.LEFT)

#起 範例
exstart_frame = tk.Frame(root)
exstart_frame.pack(side=tk.TOP)
exstart_label = tk.Label(exstart_frame, text='範例：xxxx-xx-xx')
exstart_label.pack(side=tk.LEFT)

#訖
finishTime_frame = tk.Frame(root)
finishTime_frame.pack(side=tk.TOP)
finishTime_label = tk.Label(finishTime_frame, text='起：')
finishTime_label.pack(side=tk.LEFT)
input_finishdate_var = tk.StringVar()
finishdate_widget = tk.Entry(finishTime_frame,textvariable = input_finishdate_var)
input_finishdate_get = input_startdate_var.set(input_finishdate_var.get())
finishdate_widget.pack(side = tk.LEFT)

#訖 範例
exfinish_frame = tk.Frame(root)
exfinish_frame.pack(side=tk.TOP)
exfinish_label = tk.Label(exfinish_frame, text='範例：xxxx-xx-xx')
exfinish_label.pack(side=tk.LEFT)

#end = dt.datetime.today()

def generate_excel():
    stock = str(stock_entry.get())
    start_date = input_startdate_var.get()
    finish_date = input_finishdate_var.get()
    df = web.get_data_yahoo([stock],start_date,finish_date)
    df.to_excel(r'C:01.xlsx')
    #畫圖
    close_price = df["Close"]
    close_price.plot(label = "Closed Price")
    close_price.rolling(window = 20).mean().plot(label = "20MA")
    close_price.rolling(window = 60).mean().plot(label = "60MA")
    plt.legend(loc = "best")
    plt.show()

#生成
generate_btn = tk.Button(root, text='產出 excel 並畫圖',command = generate_excel)
generate_btn.pack()

root.mainloop()
