from PyQt5.QtWidgets import *
import pandas as pd
import xlwings as xw
import sys
import yfinance as yf
import numpy as np


class CloseForm(QWidget):
    def __init__(self, name = 'CloseForm'):
        super(CloseForm,self).__init__()
        self.setWindowTitle(name)
        self.resize(200,100)   # size
        # btn 1
        self.btn_done = QPushButton(self)  
        self.btn_done.setObjectName("btn_done")  
        self.btn_done.setText("Done")
        # set layout 
        layout = QVBoxLayout()
        layout.addWidget(self.btn_done)
        self.setLayout(layout)
        # signal
        self.btn_done.clicked.connect(self.close)
        
def finish_code():
    app = QApplication(sys.argv)
    closeForm = CloseForm('Finish')
    closeForm.show()
    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window')

def consolidate_files():
    choose_files()
    completo = pd.read_csv(files[0], header = 2).dropna()
    if len(files) > 1:
        for i in range(len(files)-1): 
            archivo = pd.read_csv(files[i+1], header = 2).dropna()
            completo=pd.concat([completo,archivo])
    else:
        completo
        
    finish_code()    
    return completo
    
    


class Files(QWidget):
    def __init__(self, name = 'Files'):
        super(Files,self).__init__()
        self.setWindowTitle(name)
        self.resize(400,150)   # size
        # btn
        self.btn_chooseFile3 = QPushButton(self)  
        self.btn_chooseFile3.setObjectName("btn_chooseFile")  
        self.btn_chooseFile3.setText("[---Choose all files---]")
        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.btn_chooseFile3)
        self.setLayout(layout)
        self.btn_chooseFile3.clicked.connect(self.slot_btn_chooseFile3)
        
    

    def slot_btn_chooseFile3(self):
        global files
        files, filetype = QFileDialog.getOpenFileNames(self,  
                                    "Choose Path",  
                                    "", # start path
                                    "Excel File (*csv *.xlsx *.xls *.xlsb);;All Files (*)")  

        if files != "":
            self.close()
        elif files == "":
            return 

def choose_files():
    app = QApplication(sys.argv)
    files = Files('Historicos')
    files.show()
    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Files selected')
        
def Tickers():
    data = consolidate_files()
    print(f"Hay {(data['Ticker'].value_counts() == 31).value_counts()[1]} Tickers que estan en todos los archivos")
    tickers = pd.DataFrame(data['Ticker'].value_counts())
    ticker = tickers['Ticker'] == 31
    tickers = tickers[ticker]
    Ticker = pd.DataFrame(tickers.index)
    Ticker.rename(columns={0:'Ticker'}, inplace = True)
    for i in range(len(Ticker)):
        Ticker['Ticker'][i] = Ticker['Ticker'][i].replace("*", "" )
        Ticker['Ticker'][i] = Ticker['Ticker'][i].replace(".", "-" )
        Ticker['Ticker'][i] = Ticker['Ticker'][i].replace('MXN','POP')
        Ticker['Ticker'][i] = Ticker['Ticker'][i].replace('GFNORTEO','PP')
        Ticker['Ticker'][i] = Ticker['Ticker'][i]+".MX"
        Ticker['Ticker'][i] = Ticker['Ticker'][i].replace('POP','GFNORTEO')
        Ticker['Ticker'][i] = Ticker['Ticker'][i].replace('PP.MX','MXN=X')
        
    
    Ticker['Ticker'].sort_values
    ticker = list(Ticker['Ticker'].values)
    
    
    return ticker

def stockprices(tickers):
    end = "2022-07-29"
    start = "2020-1-31"
    stockPrices = (yf.download(tickers, start = start, end = end, interval = '1mo')["Adj Close"]).dropna()
    stockPrices.pop('MXN=X')
    stockPrices.pop('KOFUBL.MX')
    return stockPrices

def tratdatos(dfsp):
    
    datos = pd.read_csv('../NAFTRAC_20200131.csv', header = 2)
    peso_nuevo29 = datos['Peso (%)'].dropna()[30::].sum() + datos['Peso (%)'][29]
    datos['Peso (%)'].iloc[29] = peso_nuevo29
    pesos = datos['Peso (%)'][0:30]
    cap = 10000
    dfspt = dfsp.transpose()
    dfp = pd.DataFrame(dfspt.index)
    dfp.rename(columns={0:'Ticker'}, inplace = True)
    dfp['Cap'] = pesos * cap
    dfp['Prices'] = dfspt['2020-02-01'].values
    dfp['Titulos'] = dfp['Cap']//dfp['Prices']
    dfpET = pd.DataFrame([dfp.iloc[19],dfp.iloc[23]])
    dfp_preview = dfp.drop(dfp[dfp['Ticker']=='MXN=X'].index)
    dfp_new = dfp_preview.drop(dfp[dfp['Ticker']=='KOFUBL.MX'].index)
    return dfp_new

def invpasiv(dpf_nw,dfsp):
    
    pasiv =  pd.DataFrame(index=range(len(dfsp)), columns=['Timestamp', 'Capital','Rendimiento', 'Rendimiento_Acumulado'])
    for i in range(len(dfsp)):
        pasiv['Capital'][i]=(dfsp.iloc[i].values * dfp_new['Titulos']).sum()
    pasiv['Timestamp'] = dfsp.index
    pasiv_1 =  pd.DataFrame(index=range(1), columns=['Timestamp', 'Capital','Rendimiento', 'Rendimiento_Acumulado'])
    pasiv_1['Timestamp'] = "2020-01-31"
    pasiv_1['Capital'] = 1000000
    pasiv_1['Rendimiento'] = 0
    pasiv_1['Rendimiento_Acumulado'] = 0
    pasivo = pd.concat([pasiv_1,pasiv])
    pasivo['Rendimiento'] = pasivo['Capital'].pct_change()
    pasivo['Rendimiento_Acumulado'] = (pasivo['Capital'].pct_change()).cumsum()
    
    return pasivo
    