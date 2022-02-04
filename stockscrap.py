"""This File contain the code to retrive the LTP for the given stocks and
    sent it via email
   Author: Chinmay Kashikar
   Date: 22-01-2022"""
import os
import smtplib
from datetime import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup

def mainCalculation():
    """This function calculate the stock current value by fetching
    its LTP Via yahoo"""
    excel_file = pd.ExcelFile(r'Stock.xlsx')
    data = pd.read_excel(excel_file, 'Position_records', skiprows=2, usecols="C:O")
    script_name_list = data['Scrip name'].to_list()
    net_quantity_list = data['Quantity'].to_list()
    purchase_price_list = data['Purchase price'].to_list()
    trailing_stop_loss = data['Trailing stop loss'].to_list()
    invested_value_list = data['Open position value'].to_list()
    target_value = data['Target'].to_list()
    ltp = []

   # Actual Web Scrappong for price
    headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                            'AppleWebKit/537.36 (KHTML, like Gecko)'
                            'Chrome/96.0.4664.110 '
                            'Safari/537.36'}
    for i in range(len(script_name_list)):
        symbol = script_name_list[i]
        url = 'https://finance.yahoo.com/quote/'+symbol+'.NS'
        request_to_make = requests.get(url, headers=headers)
        soup = BeautifulSoup(request_to_make.text, 'html.parser')
        last_trading_stock_price = soup.find('fin-streamer', {'class':'Fw(b) Fz(36px) Mb(-4px) D(ib)'}).text
        ltp.append(last_trading_stock_price)

   # Pre process the LTP - String to int and remove the ,
    process_purchase_price = []
    for i in range(len(ltp)):
        to_find = ','
        if ltp[i].find(to_find):
            without_comma = ltp[i].replace(',', '')
            process_purchase_price.append(without_comma)
        else:
            process_purchase_price.append(without_comma)

    mapping_float = map(float, process_purchase_price)
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

    for name, quant, purchase, sloss, ltp, target, invested, current, pross in zip(script_name_list,
                                                                                   net_quantity_list,
                                                                                   purchase_price_list,
                                                                                   trailing_stop_loss,
                                                                                   new_ltp_float,
                                                                                   target_value,
                                                                                   invested_value_list,
                                                                                   current_investment,
                                                                                   profit_loss):
        sub_pp = []
        sub_pp.append("Script Name: {name}\n Quantity Purchased: {quant}"
                      "\n Purchased Price: {purchase} \n Trailing Stop Loss: {sloss}"
                      "\n Last Trading Price : {ltp} \n Target Price : {target} "
                      "\n Invested Value: {invested} \n Current Value: {current}"
                      "\n Profit OR Loss: {pross}\n\n"
                      .format(name=name, quant=quant, purchase=purchase
                              , sloss=sloss, ltp=ltp, target=target, invested=invested
                              , current=current, pross=pross))
        full_status = ''.join(sub_pp)
        list_pp.append(full_status)
    sendemail(status='\n'.join(list_pp), sum_pross=str(pross_sum),
              invested_sum=str(invested_sum), current_sum=str(curent_sum))

def sendemail(status='', sum_pross='', invested_sum='', current_sum=''):
    """This function sent email
      """
    email_address_env = os.environ.get('EMAIL_ADDRESS')
    email_password_env = os.environ.get('EMAIL_PASSWORD')
    msg = "\n{status} \n\n\n Summary:\n Total Positional Invested Value:{invested_sum}\n Total Positional Current Value:{current_sum}\nTotal Portfolio Profit/Loss: {sum_pross}".format(status=status, sum_pross=sum_pross, invested_sum=invested_sum, current_sum=current_sum)
    today = datetime.today().strftime('%d-%m-%Y')
    with smtplib.SMTP('smtp.gmail.com', 587)as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()


        smtp.login(email_address_env, email_password_env)
        subject = 'Portfolio Status '
        today = datetime.today().strftime('%d-%m-%Y')
        message_to_display = f'Subject:{subject} : {today}\n\n{msg}'
        smtp.sendmail(email_address_env, email_address_env, message_to_display)
        print('Email Sent Successfully')

if __name__ == '__main__':
    mainCalculation()
