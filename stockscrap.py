"""This File contain the code to retrive the LTP for the given stocks and
    sent it via email
   Author: Chinmay Kashikar
   Date: 22-01-2022"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
#import os
import smtplib
from datetime import datetime
def mainCalculation():
    """This function calculate the stock current value by fetching
    its LTP Via yahoo"""
    excelFile = pd.ExcelFile(r'Stock.xlsx')
    data = pd.read_excel(excelFile, 'Position_records', skiprows=2, usecols="C:O")
    script_name_list = data['Scrip name'].to_list()
    net_quantity_list = data['Quantity'].to_list()
    purchase_price_list = data['Purchase price'].to_list()
    trailing_stop_loss = data['Trailing stop loss'].to_list()
    invested_value_list = data['Open position value'].to_list()
    ltp = []

   # Actual Web Scrappong for price
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                            'AppleWebKit/537.36 (KHTML, like Gecko)'
                            'Chrome/96.0.4664.110 '
                            'Safari/537.36'}
    for i in range(len(script_name_list)):
        symbol = script_name_list[i]
        url = 'https://finance.yahoo.com/quote/'+symbol+'.NS'
        r = requests.get(url, headers=headers)
        soup = BeautifulSoup(r.text, 'html.parser')
        last_trading_stock_price = soup.find('fin-streamer', {'class':'Fw(b) Fz(36px) Mb(-4px) D(ib)'}).text
        ltp.append(last_trading_stock_price)

   # Pre process the LTP - String to int and remove the ,
    process_PP = []
    for i in range(len(ltp)):
        to_find = ','
        if ltp[i].find(to_find):
            without_comma = ltp[i].replace(',', '')
            process_PP.append(without_comma)
        else:
            process_PP.append(without_comma)

    mapping_float = map(float, process_PP)
    new_ltp_float = list(mapping_float)

   # Calculating current investment value
    current_investment = [a * b for a, b in zip(new_ltp_float, net_quantity_list)]

   #Calculating Current profit/loss
    profit_loss = [a - b for a, b in zip(current_investment, invested_value_list)]

    #Calculating overall Current profit/loss
    pross_sum = sum(profit_loss)
    invested_sum = sum(invested_value_list)
    curent_sum = sum(current_investment)



    list_pp = []

    for name, quant, purchase, sloss, ltp, invested, current, pross in zip(script_name_list,
                                                                           net_quantity_list,
                                                                           purchase_price_list,
                                                                           trailing_stop_loss,
                                                                           new_ltp_float,
                                                                           invested_value_list,
                                                                           current_investment,
                                                                           profit_loss):
        sub_pp = []
        sub_pp.append("Script Name: {name}\n Quantity Purchased: {quant}"
                      "\n Purchased Price: {purchase} \n Trailing Stop Loss: {sloss}"
                      "\n Last Trading Price : {ltp} \n Invested Value: {invested}"
                      "\n Current Value: {current} \n Profit OR Loss: {pross}\n\n"
                      .format(name=name, quant=quant, purchase=purchase
                              , sloss=sloss, ltp=ltp, invested=invested
                              , current=current, pross=pross))
        full_status = ''.join(sub_pp)
        list_pp.append(full_status)
    sendEmail(status='\n'.join(list_pp), sum_pross=str(pross_sum),
              invested_sum=str(invested_sum), current_sum=str(curent_sum))

def sendEmail(status='', sum_pross='', invested_sum='', current_sum=''):
    """This function sent email
      """
    EMAIL_ADDRESS = 'watchmyportfolio1607@gmail.com' #os.environ.get('EMAIL_ID')
    EMAIL_PASSWORD = 'rifkynungjlzucaj' #os.environ.get('Email_Password')
    msg = "\n{status} \n\n\n Summary:\n Total Positional Invested Value:{invested_sum}\n Total Positional Current Value:{current_sum}\nTotal Portfolio Profit/Loss: {sum_pross}".format(status=status, sum_pross=sum_pross, invested_sum=invested_sum, current_sum=current_sum)
    today = datetime.today().strftime('%d-%m-%Y')
    with smtplib.SMTP('smtp.gmail.com', 587)as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()


        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        subject = 'Portfolio Status '
        today = datetime.today().strftime('%d-%m-%Y')
        msg1 = f'Subject:{subject} : {today}\n\n{msg}'
        smtp.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg1)
        print('Email Sent Successfully')

if __name__ == '__main__':
    mainCalculation()

