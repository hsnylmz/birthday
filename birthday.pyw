import mysql.connector
import datetime
import win32com.client as win32
import os

from datetime import date
from datetime import datetime
from datetime import timedelta
signatureimage = "C:/birthday/card.jpg"

pas="5555555555555555555555555555"
usr="root"
hst="localhost"

db = mysql.connector.connect( user=usr, password=pas, host=hst)
cur = db.cursor()

SQL = 'USE b_day;'
cur.execute(SQL)

SQL = 'SELECT count(adi_soyadi) as tane FROM bday WHERE DAY(dogum_tarihi) = DAY(CURRENT_DATE) AND MONTH(dogum_tarihi)=MONTH(CURRENT_DATE)'
cur.execute(SQL)
results = cur.fetchall()

SQL_2 = 'SELECT adi_soyadi, dogum_tarihi,eposta FROM bday WHERE DAY(dogum_tarihi) = DAY(CURRENT_DATE) AND MONTH(dogum_tarihi)=MONTH(CURRENT_DATE)'
cur.execute(SQL_2)
results_2 = cur.fetchall()

for tane in results:
    if(tane[0]>0):
        for kisi in results_2:
            ad=kisi[0]
            eposta=kisi[2]

            print(ad)
            print(eposta)


            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Doğum Gününüz Kutlu Olsun.'
            mailItem.BodyFormat = 1
            mailItem.Body = 'Hasan YILMAZ '

            attachment = mailItem.Attachments.Add(signatureimage)
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "tebrik")
            mailItem.HTMLBody = "<html><body><img src=""cid:tebrik""></body></html>"

            mailItem.To = eposta
            

            #mailItem.CC = email
            #mailItem.BCC = email

            
            #mailItem.SenderEmailAddress = email
            #mailItem.SentOnBehalfOfName = email


            #mailItem.Display()
            mailItem.Save()
            mailItem.Send()

