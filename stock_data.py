import numpy as np
import pandas_datareader.data  as web
import matplotlib.pyplot as plt
import datetime as dt
import pandas as pd
import csv
import xlrd
import scipy.optimize as sco
start = dt.datetime(2016,1,1)
end = dt.datetime(2016,12,31)
#end = dt.datetime.today()
book = xlrd.open_workbook(r'C:\学校文件\天津大学金融工程\金融工程\投资策略\\'+'stock_poor_sh.xlsx')
sheet_1 = book.sheet_by_name('Sheet2')
stock_name = sheet_1.col(0)
stock_sum = len(stock_name)
symbols = pd.DataFrame()
#stock_name = pd.DataFrame(stock_name)
for m in range(stock_sum):
	symbols = symbols.append([stock_name[m].value])
#print(symbols)#这里xlsx的表格有个text:,因为是cell，text：表示属性，所以要加value才是后面的值
noa =len(symbols)
#print(symbols.iat[1,0])
#print(symbols.iat[1446,0])
data =pd.DataFrame()
X = pd.DataFrame()
#symbols.iat[num,1]
stock_num = 0
#writer = csv.writer(outputfile)
#outputfile.write("num#stock_data\n")
for num in range(1,3):
	print('Get Stock Info:no.'+str(num))
	try:
		data[symbols.iat[num,0]] = web.DataReader(symbols.iat[num,0],'yahoo',start,end)['Close']
		X = X.append(data[symbols.iat[num,0]])
		data_name = 'df'+str(num)
		data_name = data[symbols.iat[num,0]]
		#if num == 1:
			#data[symbols.iat[num,0]] .to_excel('stock_data_3.xlsx',startcol=0,sheet_name='stock',index = True)
			#data_name.to_excel('stock_data_3.xlsx', startcol=0, sheet_name='stock', index=True)
			#data_1 = data[symbols.iat[num,0]]
		#else:
			#data[symbols.iat[num,0]].to_excel('stock_data_3.xlsx',startcol=4,sheet_name='stock',index=False)
			#data_name.to_excel('stock_data_3.xlsx', startcol=4, sheet_name='stock', index=True)
			#X = [data_1,data_name]
			#Z = pd.concat(X)
	except:
		pass
	continue
X.to_excel('stock_data_3.xlsx',index=True)
	#data.columns = symbols
#num = len(data[symbols[0]])
#plt.figure(figsize=(8,6))
#plt.grid(True)
#p1 = plt.plot(data[symbols[0]],label='MSFT')
#p2 = plt.plot(data['AMZN'],label='AMZN')
#p3 = plt.plot(data['DB'],label='DB')
#p4 = plt.plot(data['GLD'],label='GLD')
#p5 = plt.plot(data['000002.SZ'],label='VANKE')
#plt.legend()
#plt.show()
