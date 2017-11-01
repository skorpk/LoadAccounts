'''
Created on 21 марта 2016 г.
1.Get data from SQL Server. 
2.Create xml file like TKR3416*
3.The data will be sended into FFOMS.
@author: SKrainov
'''
sqlConnStr = ('DRIVER={SQL Server};Server=srv-cnt-db2;Database=AccountOMS;'+
            'Trusted_Connection=YES') 


import pyodbc #to connect to SQL Server
import os
import zipfile

def createZIP_DelXML(pathName,FileName):
    zipf = zipfile.ZipFile(os.path.join(pathName, FileName+".oms"), 'w',zipfile.ZIP_DEFLATED)
    zipf.write(pathName+'/'+FileName+".xml",FileName+'.xml')        
    zipf.close()
    os.remove(pathName+'/'+FileName+".xml")
''' Insert information about file'''
def insertIntoSendingFile(fileNameXML,reportMonth,reportYear,code):
    sqlConn = pyodbc.connect(sqlConnStr, autocommit = True)
    curs = sqlConn.cursor()
    print(fileNameXML)
    curs.execute("EXEC dbo.usp_InsertSendingInformationAboutFile @nameFile =?, @reportMonth =?, @reportYear =?, @code = ?",fileNameXML,reportMonth,reportYear,code)
           

def getXML(nameFile,reportMonth,reportYear,code,pathName):    
    fileNameXML=fName+('000'+str(code))[-4:]
    insertIntoSendingFile(nameFile,reportMonth,reportYear,code)
    
    sqlConn = pyodbc.connect(sqlConnStr, autocommit = True)
    curs = sqlConn.cursor()
    curs.execute("EXEC dbo.usp_GetXMLSendingDataToFFOMS @nameFile=?,@reportMonth=?,@reportYear=?,@code=? ",nameFile,reportMonth,reportYear,code)     
    for r in curs:
        file=open(os.path.join(pathName,fileNameXML+".xml"), mode='w')      
        file.write('<?xml version="1.0" encoding="Windows-1251"?>')
        file.write(r.colXML)
        file.close()    
        '''createZIP_DelXML(pathName,fileNameXML)'''  
        
        
    
            
sqlConn = pyodbc.connect(sqlConnStr, autocommit = True)
dirName=r"f:\test\TKRFiles"
'''
Данные передаются пользователем
в reportMM ставим отчетный месяц
'''
reportMM=9
reportYYYY=2017
fName='TKR34'+(str(reportYYYY)[-2:])
'''
sqlConn = pyodbc.connect(sqlConnStr, autocommit = True)
curs = sqlConn.cursor()
curs.execute("SELECT ISNULL(MAX(NumberOfEndFile),0)+1 as Number FROM dbo.t_SendingFileToFFOMS WHERE ReportYear=?", reportYYYY)
for row in curs:
    print(fName+('000'+str(row.Number))[-4:])
    getXML(fName,reportMM,reportYYYY,row.Number,dirName)            
    print('File unloaded')    
    reportMM=reportMM+1
   ''' 
    

