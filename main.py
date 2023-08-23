import numpy as np
import pandas as pd
import yfinance as yf
import smtplib
from email.message import EmailMessage
import datetime

start_date='2022-08-23' #Start Date of the Portfolio
today=str(datetime.date.today()+datetime.timedelta(days=1)) #Today's date

email_id='myemail@gmail.com' #Enter your Gmail ID
sender_name='Myself' #Enter your name
app_password='abcdefghijklm' #Generate App password and paste here

tickers=['RELIANCE.NS','ADANIENT.NS','LT.NS','AXISBANK.NS','ITC.NS'] #Enter your tickers

quantities={'RELIANCE.NS':4, #Enter Ticker quantites
            'ADANIENT.NS':3,
            'LT.NS':6,
            'AXISBANK.NS':14,
            'ITC.NS':31}

buy_avg={'RELIANCE.NS':2366.45, #Enter Ticker Buy avg. prices
         'ADANIENT.NS':3030.65,
         'LT.NS':1877.05,
         'AXISBANK.NS':747.05,
         'ITC.NS':315.90}

invested_value=[] #Setting empty list for invested value (incase qty/buy avg price needs to be changed)
final_df={} #Initializing daily stock report 

for ticker in tickers: #Looping through Tickers

    stock_df=yf.download(ticker,start=start_date,end=today) #Downloading the Ticker from Yahoo Finance API
    stock_df['Change']=stock_df['Adj Close'].diff() #Calculating price change
    stock_df['%Change']=stock_df['Adj Close'].pct_change()*100 #Calculating % change
    sub_df=stock_df.iloc[-1] #Selecting last row i.e today's date
    final_df[ticker]=sub_df.to_dict() #Converting that row into a dictionary and adding it to final_df
    invested_value.append((quantities[ticker]*buy_avg[ticker])) #Calculating Invested values

portfolio_df=pd.DataFrame(final_df).T #Dataframing the final_df dictionary and transposing it
portfolio_df.index.name='Tickers' #Setting index name as 'Tickers'

#Creating necessary columns
portfolio_df['Quantity']=quantities 
portfolio_df['Invested Amount']=invested_value
portfolio_df['Market Value']=portfolio_df['Close']*portfolio_df['Quantity']
portfolio_df['Days P&L']=portfolio_df['Change']*portfolio_df['Quantity']
portfolio_df['All time P&L']=portfolio_df['Market Value']-portfolio_df['Invested Amount']
portfolio_df['All time P&L(%)']=(((portfolio_df['Market Value']-portfolio_df['Invested Amount'])/portfolio_df['Invested Amount'])*100)
portfolio_df=portfolio_df.round(2) #Rounding all values to 100th place

#Summing the Portfolio values to get total change/stats
fund_df=pd.DataFrame()
fund_df.index=[stock_df.index[-1]] #Setting index as the last date (today) of the stock dataframe
fund_df['Invested Amount']=[portfolio_df['Invested Amount'].sum()]
fund_df['Market Value']=[portfolio_df['Market Value'].sum()]
fund_df['Days P&L']=[portfolio_df['Days P&L'].sum()]
fund_df['Daily percent change']=((fund_df['Days P&L'])/(fund_df['Market Value']-fund_df['Days P&L']))*100
fund_df['All Time percent change']=((fund_df['Market Value']-fund_df['Invested Amount'])/fund_df['Invested Amount'])*100

#Read the output excel file
excel_df=pd.read_excel('FundAutomation.xlsx',index_col=0)

#Function to concatenate the fund dataframe to existing excelsheet
def service(send_email=True):
    concat_df=pd.concat([excel_df,fund_df],ignore_index=False)
    writer = pd.ExcelWriter('FundAutomation.xlsx', engine='xlsxwriter')
    concat_df.to_excel(writer,sheet_name='Portfolio Movement',index=True)
    writer.close()
    
    if send_email: #Sends email about today's tickers movement
        msg = EmailMessage()
        msg['Subject'] = f'Ticker Summary for {today}'
        msg['From'] = f"{sender_name} <{email_id}>"
        msg['To'] = f"{email_id}"
        msg.add_alternative(portfolio_df.to_html(index=True), subtype='html')

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(email_id, app_password) 
            smtp.send_message(msg)

if len(excel_df.columns)==0: #If the excelsheet is empty call the function
    service()
elif excel_df.index[-1]==fund_df.index: #If the latest index (date) of the excel sheet is equal to the date in the fund data frame, pass (as the dates will get repeated)
    pass
else: #Else call function
    service()




