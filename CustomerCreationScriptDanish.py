# -*- coding: utf-8 -*-
"""
Created on Sat Jan 25 10:51:54 2020

@author: thera
"""

import pyodbc
from datetime import date
import shutil, os
import xlwings as xw

# This core was written as an automation script for a company's initial customer process. Creating the customer in the database, creating a casefolder, copying over standardized forms, 
# and inputting some of the customer data into a specified Excel worksheet. Questions and specifics have been removed, for security's sake.


#Forbind til databasen
dbcon = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=<insert database>;')
cursor = dbcon.cursor()

#Spørg efter kundeinformation
#today = date.today()
#cdate = today.strftime("%d/%m/%Y")
cname = input("<customername>")
caddress = input("<customeraddress>")
czip = input("<customerzip>")
ccity = input("<customercity>")
cphone = input("<customerphone>")
cemail = input("<customeremail>")
cnotes = input("<customernotes>")
cmount = input("<customermount>") #spørger om der er en anden monteringsadresse
cdeliv = input("<customerdeliv>") #spørger om der er en anden leveringsadresse
cmadress = ''
cmzip = ''
cmcity = ''
cladress = ''
clzip = ''
clcity = ''

#Funktion for mulig anderledes adresse, spørger efter ændringer
if cmount == "ja":
   cmadress = input("<customermountadress>")
   cmzip = input("<customermountzip>")
   cmcity = input("<customermountcity>")

if cdeliv == "ja":
    cladress = input("<customerdelivadress>")
    clzip = input("<customerdelivzip>")
    clcity = input("<customerdelivcity>")

if cmount == 'nej' or '':
    cmadress = caddress
    cmzip = czip
    cmcity = ccity

if cdeliv == 'nej' or '':
    cladress = caddress
    clzip = czip
    clcity = ccity
        
#Input de forskellige variabler i databasen
print("Opretter kunden i databasen...")
cinfo = (cname, caddress, czip, ccity, cphone, cemail, cmadress, cmzip, cmcity, cladress, clzip, clcity, cnotes)
cursor.execute("INSERT INTO <database> (Navn, Adresse, Postnummer, [By], Tlfnr, Email, MAdresse, MPostnummer, [MBy], LAdresse, LPostnummer, [LBy], Noter) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", cinfo)
dbcon.commit()

#Hent kundenr fra databasen
print("Henter kundens kundenr...")
kundenr = cursor.execute("select kundenr from kunder where Tlfnr = ?", [cphone]).fetchval()

#Lav en mappe i salgsmappen med kundenr, navn og adresse
print("Opretter kundemappe...")
path = "<folder>/" + str(kundenr) + " " + cname + " - " + caddress + ", " + czip + " " + ccity

#Kopier alt fra Formularer mappen ind i den nye sagsmappe
print("Overfører standardblanketter til kundemappen...")
shutil.copytree("Formularer", path)

#Indtast kunden i regnearket
print("Indtaster kunden i regnearket")

os.chdir(path)
retval = os.getcwd()
print(retval)

app = xw.App(visible=False)
book = app.books.open('<excelfile>')
sht = book.sheets('<excelsheet>')
sht.range('C5').value = cname
sht.range('C6').value = caddress
sht.range('C7').value = czip
sht.range('E7').value = ccity
sht.range('C8').value = cmadress
sht.range('C9').value = cmzip
sht.range('E9').value = cmcity
sht.range('C10').value = cladress
sht.range('C11').value = clzip
sht.range('E11').value = clcity
sht.range('C12').value = cemail
sht.range('C13').value = cphone
book.save('<excelfile1>')
app.kill()
os.unlink('<excelfile>')

os.chdir("..")
os.chdir("..")
os.startfile("<root>" + path)
retval2 = os.getcwd()
print(retval2)

print("<goodbye note>")
input("Tryk Enter for at afslutte")