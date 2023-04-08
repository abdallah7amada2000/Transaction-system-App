import pandas

# sheet_data=input("Enter sheet data")
#data OldOrdersCallCenterData
data = pandas.read_excel("OldOrdersCallCenterData (25).xlsx")
data.drop([1,2,3,4,5,],axis=0, inplace=True)
res = {'Order Management System': 'ID', 'Unnamed: 1': 'Order Num', 'Unnamed: 2': 'Account', 'Unnamed: 3': 'Rest', 'Unnamed: 4': 'Total Price', 'Unnamed: 5': "nan", 'Unnamed: 6': 'Status', 'Unnamed: 7': 'Cancel Reason', 'Unnamed: 8': 'Order Source', 'Unnamed: 9': 'Discount Type', 'Unnamed: 10': 'Discount', 'Unnamed: 11': 'Notes', 'Unnamed: 12': 'Hidden Comment', 'Unnamed: 13': 'Created', 'Unnamed: 14': 'Created By', 'Unnamed: 15': 'Updated', 'Unnamed: 16': 'Updated By', 'Unnamed: 17': 
'New', 'Unnamed: 18': 'Assembling', 'Unnamed: 19': 'Prepared', 'Unnamed: 20': 'Outed', 'Unnamed: 21': 'Total Time', 'Unnamed: 22': 'Phones', 'Unnamed: 23': 'Geo Area'}
data.rename(columns=res,inplace=True)
data.drop([0],inplace=True,axis=0)
data.drop(['ID', 'Order Num', 'Account', 'Rest', 'nan',
                'Cancel Reason', 'Order Source', 'Discount Type', 'Discount', 'Notes',
                'Hidden Comment', 'Created',  'Updated', 'Updated By',
                'New','Total Price', 'Assembling', 'Prepared', 'Outed', 'Total Time', 'Phones','Geo Area'],
          axis=1,
          inplace=True
          )

ex= data[data["Status"] == "Delivered" ] 
ex2 = data[data["Status"] == "Canceled" ]
data = pandas.concat([ex,ex2])
names = data['Created By']
def countduplicate(name):
    return len(data[data['Created By'].str.fullmatch(name)])

n = names.apply(countduplicate)
data['Created By'].drop_duplicates(inplace=True)
data.drop('Status',axis=1, inplace=True)
data.insert(1,'TC',n)
data.drop_duplicates("Created By", inplace=True)
# print(data)
# data.to_excel("52.xlsx",index=False)





#data call status
data1=pandas.read_excel('tran.xlsx')
df2 = data1.pivot_table(index=['TYPE'], aggfunc ='size')
df2.drop(columns=['TYPE'] , inplace=True )
print(df2)
# df2.to_excel("44.xlsx")





#_data_complaint_Sheet
data2=pandas.read_excel("Complaints (4).xlsx")
data2.drop([0,1,2],axis='index'
, inplace=True)
data0 = data2.rename(columns=data2.iloc[0], inplace=True)
data2.drop([3], axis="index",inplace=True)
df3 = data2.pivot_table(index=['Created By'], aggfunc ='size')
# print(df3)
# data2.to_excel('55s.xlsx')


# df_finel = data.add(df2, fill_value=0)
# print(df_finel)