import pandas as pd
import datetime as dt

df = pd.read_excel(r"C:\Users\mxr29\Desktop\Fall 2020\Metal Works\Data\MIM.xlsx", sheet_name= 'Ticket System')
df.dropna(axis = 1, how = 'all', inplace= True)


df = df.drop(df.columns[[0,1,2,4,6,8,9,10,12,13,14,15,17,18,19]], axis = 1)
df = df.iloc[9:,:]


df.columns = ['Ticket ID','Ticket Date','Company','Container #','Area','Weight']
df['Area'] = df['Area'].shift(axis = 0, periods=-1)
df['Weight'] = df['Weight'].shift(axis = 0 , periods=-1)
df.dropna(axis = 0, how = 'all', inplace= True)


df['Ticket ID'] = df['Ticket ID'].fillna(method='ffill',axis = 0 )
df['Ticket Date'] = df['Ticket Date'].fillna(method='ffill',axis = 0 )
df['Company'] = df['Company'].fillna(method='ffill',axis = 0 )



df['Ticket Date'] = df['Ticket Date'].dt.strftime('%m/%d/%Y')


for i in range (len(df)):
    if df['Ticket ID'].iloc[i-1] > df['Ticket ID'].iloc[i]:
        split = i



receiving_df = df.iloc[:split,:]
shipping_df = df.iloc[split:,:]



writer = pd.ExcelWriter(r"C:\Users\mxr29\Desktop\Fall 2020\Metal Works\Data\MIM_out.xlsx", engine='xlsxwriter')
receiving_df.to_excel(writer, sheet_name ="Receiving",index=False)
shipping_df.to_excel(writer, sheet_name ="Shipping",index=False)
writer.save()