from datetime import datetime
import pandas as pd
import os
import sys
from PyQt5.QtWidgets import QApplication, QDialog, QMainWindow, QTableWidgetItem, QFileDialog
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from PyQt5.uic import loadUi

class CryptoTrade:
    def __init__(self):
        self.trade_history = []
        self.report = pd.DataFrame(columns=['Date', 'Type', 'Coin', 'Amount', 'Price', 'Total'])
        self.portfolio = pd.DataFrame(columns=['Coin', 'Amount'])
        
    def buy(self, coin, amount, price):
        total = amount * price
        self.trade_history.append((datetime.now(), 'BUY', coin, amount, price, total))
        self.report.loc[len(self.report)] = [datetime.now(), 'BUY', coin, amount, price, total]
        if coin not in self.portfolio['Coin'].values:
            self.portfolio.loc[len(self.portfolio)] = [coin, amount]
        else:
            index = self.portfolio.index[self.portfolio['Coin'] == coin][0]
            self.portfolio.loc[index, 'Amount'] += amount

    def sell(self, coin, amount, price):
        total = amount * price
           
        self.trade_history.append((datetime.now(), 'SELL', coin, amount, price, total))
        self.report.loc[len(self.report)] = [datetime.now(), 'SELL', coin, amount, price, total]
        index = self.portfolio.index[self.portfolio['Coin'] == coin][0]
        self.portfolio.loc[index, 'Amount'] -= amount
        if self.portfolio.loc[index, 'Amount'] == 0:
            self.portfolio = self.portfolio.drop(index)
        elif self.portfolio.loc[index, 'Amount'] < 0:
            raise ValueError("Selling more coins than owned in the portfolio")
            
    def get_report(self):
        return self.report
    
    def get_portfolio(self):
        return self.portfolio 
    
    def save_to_excel(self, filename):
        writer = pd.ExcelWriter(filename)
        self.report.to_excel(writer, sheet_name='Trade History', index=False)
        self.portfolio.to_excel(writer, sheet_name='Portfolio', index=False)
        writer.save()
        
    def load_from_excel(self, filename):
        if not os.path.exists(filename):
            return
        self.report = pd.read_excel(filename, sheet_name='Trade History')
        self.portfolio = pd.read_excel(filename, sheet_name='Portfolio') 

class CryptoTradeUI(QMainWindow):
    def __init__(self):
        super(CryptoTradeUI, self).__init__()
        loadUi('Crypto_Trade_UI.ui', self)
        self.trade = CryptoTrade()
        self.populate_table(self.trade.get_report())
        self.populate_portfolio(self.trade.get_portfolio())
        self.trade_type.addItem('BUY')
        self.trade_type.addItem('SELL')
        self.coin_symbols = ['BTC', 'ETH', 'XRP', 'LTC', 'BCH', 'ADA']
        self.coin_symbol.addItems(self.coin_symbols)
        self.submit_trade_button.clicked.connect(self.submit_trade)
        self.save_button.clicked.connect(self.save_to_excel)
        self.load_button.clicked.connect(self.load_from_excel)
        
    def populate_table(self, report):
        self.trade_history.setRowCount(0)
        for i, row in report.iterrows():
            self.trade_history.insertRow(i)
            for j in range(len(row)):
                self.trade_history.setItem(i, j, QTableWidgetItem(str(row[j])))

    def populate_portfolio(self, portfolio):
        self.portfolio_table.setRowCount(0)
        for i, row in portfolio.iterrows():
            self.portfolio_table.insertRow(i)
