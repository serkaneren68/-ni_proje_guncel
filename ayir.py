
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl import Workbook

r = open("toplu.html", "r")


workbook = Workbook()
sheet = workbook.active
source = BeautifulSoup(r,"html.parser")

sonra = source.find_all("font",attrs={"color":"#CC0000"})
orta = source.find_all("td",attrs={"class":"dt-center vcenter"})
once =  source.find_all("strong")
ensonra = source.find_all("td",attrs={"class":"dt-center vcenter sorting_1"})
k =0
j = 1
for i in once:

    if(k%2 == 0):
        print(i.text)
        sheet["A{}".format(j)].value = i.text
        j=j+1
    k=k+1
k =0
j = 1


for i in sonra:
    if ( ")" in i.text)-1:
        sheet["B{}".format(j)].value = i.text
        j=j+1
k =0
j = 1

for i in sonra:
    if ( ")" in i.text):
        if("İngilizce-Fransızca" in i.text):
            sheet["M{}".format(j)].value = "(İngilizce-Fransızca)"
        elif("KKTC Uyruklu" in i.text):
            sheet["M{}".format(j)].value = "(KKTC Uyruklu)"
        elif("Almanca" in i.text):
            sheet["M{}".format(j)].value = "(Almanca)"
        elif("Fransızca" in i.text):
            sheet["M{}".format(j)].value = "(Fransızca)"
        elif("İngilizce" in i.text):
            sheet["M{}".format(j)].value = "(İngilizce)"
        elif("İtalyanca" in i.text):
            sheet["M{}".format(j)].value = "(İtalyanca)"
        elif("İspanyolca" in i.text):
            sheet["M{}".format(j)].value = "(İspanyolca)"
        elif("Rusça" in i.text):
            sheet["M{}".format(j)].value = "(Rusça)"
        elif("Arapça" in i.text):
            sheet["M{}".format(j)].value = "(Arapça)"

        else:
            sheet["M{}".format(j)].value = " "
        j=j+1
k =0
j = 1

for i in once:

    if(k%2 == 1):
        print(i.text)
        sheet["C{}".format(j)].value = i.text
        j=j+1
    k=k+1

k =0
j = 1
for i in orta:
    print(i.text)
    if(k%8==0):
        sheet["K{}".format(j)].value = i.text
    elif(k%8==1):
        sheet["D{}".format(j)].value = i.text
    elif(k%8==2):
        sheet["E{}".format(j)].value = i.text
    elif(k%8==3):
        sheet["F{}".format(j)].value = i.text
    elif(k%8==4):
        sheet["G{}".format(j)].value = i.text
    elif(k%8==5):
        sheet["H{}".format(j)].value = i.text
    elif(k%8==6):
        sheet["I{}".format(j)].value = i.text
    elif(k%8==7):
        sheet["J{}".format(j)].value = i.text
        j=j+1
    k=k+1
j = 1
for i in ensonra:
    sheet["L{}".format(j)].value = i.text
    j=j+1

workbook.save("toplu2.xlsx")
