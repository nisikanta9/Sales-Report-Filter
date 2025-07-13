# (Sales Report Filter)
# (It will automate our sales Excel sales filtering tasks using python-save time,get clean and filtered report fast)

import pandas as pd
df=pd.read_excel('babu.xlsx',sheet_name="Summary",parse_dates=["Date"])
df.columns=df.columns.str.strip()
min_sales=float(input('Enter minimum sales amount:'))
product=input('Enter product:').strip().upper()
region=input('enter region:').strip().capitalize() 
start_date=input('Enter start date(YYYY-MM-DD):')
end_date=input('Enter end date(YYYY-MM-DD):')
start=pd.to_datetime(start_date)
end=pd.to_datetime(end_date)
filtered=df[(df['Date']>=start)&(df['Date']<=end)&(df['Region']==region)&(df['Product']==product)&(df['Sales']>min_sales)]
filtered.to_excel("output.xlsx",index=False)
print('Data filtered & saved to output.xlsx')