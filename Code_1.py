#code for projerct
import pandas as pd
from datetime import date
from datetime import timedelta


today = date.today()
#print(today)
database=pd.read_excel(r'C:\Users\HP\Desktop\Sem3\ITW\Project.xlsx')
#print(database)

Name = database['Name']
Rack = database['Rack']
Shelf = database['Shelf']
Expiry = database['Exp_Date']
print(Expiry)
print('')

_Today = pd.to_datetime(today)
_a_day = pd.to_datetime((date.today()+timedelta(days=1)))
_a_week = pd.to_datetime((date.today()+timedelta(days=7)))
_15_days = pd.to_datetime((date.today()+timedelta(days=15)))
_Expired = pd.to_datetime(today)



a_day_database = pd.DataFrame(columns = ['Name','Rack','Shelf']) 
a_week_database = pd.DataFrame(columns = ['Name','Rack','Shelf'])
_15_days_database = pd.DataFrame(columns = ['Name','Rack','Shelf'])
expired_database = pd.DataFrame(columns = ['Name','Rack','Shelf'])


Total_Item = len(database)-1

i=0

while(i<Total_Item):
    if(Expiry[i]<_Today):
        expired_database = expired_database.append({'Name': Name[i],'Rack': Rack[i],'Shelf': Shelf[i]}, ignore_index=True)
    elif(_Today < Expiry[i]<=_a_day):
         a_day_database = a_day_database.append({'Name': Name[i],'Rack': Rack[i],'Shelf': Shelf[i]}, ignore_index=True)
    elif(_Today < Expiry[i]<=_a_week):
        a_week_database = a_week_database.append({'Name': Name[i],'Rack': Rack[i],'Shelf': Shelf[i]}, ignore_index=True)
    elif(_Today < Expiry[i]<=_15_days):
        _15_days_database = _15_days_database.append({'Name': Name[i],'Rack': Rack[i],'Shelf': Shelf[i]}, ignore_index=True)
    i=i+1
    
print('Expired:- ')
print(expired_database)
print('')
print('In a day:- ')
print(a_day_database)
print('')
print('In a week:- ')
print(a_week_database)
print('')
print('In 15 days:- ')
print(_15_days_database)
print('')    


writer = pd.ExcelWriter('Project.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
expired_database.to_excel(writer, sheet_name='Sheet4')
a_day_database.to_excel(writer, sheet_name='Sheet2')
a_week_database.to_excel(writer, sheet_name='Sheet3')
_15_days_database.to_excel(writer, sheet_name='Sheet3')

# Close the Pandas Excel writer and output the Excel file.
writer.save()
#expired_database.to_excel(r'C:\Users\HP\Desktop\Sem3\ITW\Project_output.xlsx', sheet_name='Already expired', index = False)
#a_day_database.to_excel(r'C:\Users\HP\Desktop\Sem3\ITW\Project_output.xlsx', sheet_name='A day', index = False)
#a_week_database.to_excel(r'C:\Users\HP\Desktop\Sem3\ITW\Project_output.xlsx', sheet_name='A Week', index = False)
#_15_days_database.to_excel(r'C:\Users\HP\Desktop\Sem3\ITW\Project_output.xlsx', sheet_name='15 days', index = False)
