
import pyodbc #to connect to SQL Server
import os
import xlrd
import math
from numpy.ma.core import getdata

sqlConnStr = ('DRIVER={SQL Server};Server=srv-cnt-db2;Database=AccountOMS;Trusted_Connection=YES')
rb = xlrd.open_workbook('f:/AdditionalAccounts/2017/AMP_09_2017.xls',formatting_info=True) 

def insertData(sqlCommand):
    sqlConn = pyodbc.connect(sqlConnStr, autocommit = True)
    curs = sqlConn.cursor()
    curs.execute(sqlCommand)
    
def getData():    
    sheet=rb.sheet_by_index(0)
    i=1
    accumulator = ('''INSERT dbo.t_AdditionalAccounts(CodeLPU ,CodeSMO ,ReportMonth ,ReportYear ,NumberRegister ,Letter ,DateRegistration ,DateAccount ,AmountPayment)
    SELECT CodeLPU ,CodeSMO ,ReportMonth ,ReportYear ,NumberRegister ,Letter ,DATEADD(DAY,-2,cast(v.Dateregistration AS datetime)),DATEADD(DAY,-2,CAST(cast(v.DateAccount AS datetime) AS DATE)),AmountPayment
    FROM (values''')
    for rownum in range(sheet.nrows):
        while i<rownum:
            s='(\''+str(int(sheet.cell(i,0).value))+'\',\''+str(int(sheet.cell(i,3).value))+'\','+str(int(sheet.cell(i,4).value))+',2017,'+str(int(sheet.cell(i,5).value))+',\'P\','+str(int(sheet.cell(i,7).value))+','+str(int(sheet.cell(i,6).value))+','+str(sheet.cell(i,2).value)+'),'        
            accumulator+=s
            i+=1
    accumulator+=s[:-1]+') v(CodeLPU,CodeSMO,ReportMonth,ReportYear,NumberRegister,Letter,Dateregistration,DateAccount,AmountPayment)'
    'insertData(accumulator)'
    print(accumulator)

  
getData()
rb = xlrd.open_workbook('f:/AdditionalAccounts/2017/SMP_09_2017.xls',formatting_info=True) 
getData()

