#!/usr/bin/python3

from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter

addr_lst =[]
zip_lst = []
city_lst = []
phone_lst = []
fname_lst = []
lname_lst = []


workbook = xlsxwriter.Workbook('output-list.xlsx')
worksheet  = workbook.add_worksheet()


worksheet.write('A1','Mr/s')
worksheet.write('B1','Job Titel')
worksheet.write('C1','Firstname')
worksheet.write('D1','Lastname')
worksheet.write('E1','Street')
worksheet.write('F1','Zip code')
worksheet.write('G1','City')
worksheet.write('H1','Phone')


for j in range(1,28):
    html = urlopen("https://www.herold.at/telefonbuch/j38Bh_hirtenberg/?page=" + str(j))


    bsobj = BeautifulSoup(html.read(),'html.parser')

###################################################################
    address = bsobj.find_all('p',class_='address')
    tel = bsobj.find_all('div',class_='col-lg-7 d-none d-sm-block')
    name = bsobj.find_all('h2')
###################################################################

    for i in range(len(address)):
        strtel = ""
        stradd = ""
        city = ""
        strname = ""

############    Addressa    ###########
        stradd = address[i].get_text()
        stradd_comma = stradd.split(',')
        if(len(stradd_comma) == 1):
            continue
############    Addressa End    #######




############    Name    ###############
        strname = list(name[i].a.children)[0]
        strname_lst = (strname.get_text().split(' '))
        if(len(strname_lst) == 1):
            continue    
############    Name End    ###########




############    ZIP     ###########
        stradd_lst = stradd.split(' ')
        zip = ''
        for g in range(len(stradd_lst)):
            if(len(stradd_lst[g]) == 4 and stradd_lst[g].isnumeric()):
                zip = stradd_lst[g]
        if(zip == ''):
            continue
############    ZIP END     ######





############    CITY     ##########
        for q in range(stradd_lst.index(zip)+1,len(stradd_lst)):
            city += stradd_lst[q] + " "
############    CITY End     ######





############    TEL     ##########
        if(len(tel[i].findChildren()) <= 1):
             continue
        else:
            strtel = list(tel[i].div.div.div.div)[1]
############    TEL End     ######






###### Remove spaces at the beginning and end of the str-address#########
        if(len(stradd_comma) > 2):
            addr = stradd_comma[1][1:]
        else:
            if(stradd_comma[0][len(stradd_comma[0])-1] == ' '):
                addr = stradd_comma[0][:-1]
            else:
                addr = stradd_comma[0]
###### Remove spaces at the beginning and end of the str-address   END ###







##############Check if tel  starts with 6 ##########################
        strtel_lst = strtel.replace(" ","")
        phone = strtel_lst[1:].replace('-','')
        if(phone[0] != '6'):
            continue
##############Check if tel number starts with 6 END ##########################





###### Check if client has 3 or more names  ##############################
        if(len(strname_lst) < 3):
            fname = strname_lst[1]
            lname = strname_lst[0]
        else:
            fname = "**" + strname_lst[1]
            lname = strname_lst[0]
###### Check if client has 3 or more names END ############################





        fname_lst.append(fname)
        lname_lst.append(lname)
        phone_lst.append(int(phone))
        addr_lst.append(addr)
        zip_lst.append(int(zip))
        city_lst.append(city[:-1])


for p in range(len(fname_lst)):
    worksheet.write('C' + str(p+2),fname_lst[p])
    worksheet.write('D' + str(p+2),lname_lst[p])
    worksheet.write('E' + str(p+2),addr_lst[p])
    worksheet.write('F' + str(p+2),zip_lst[p])
    worksheet.write('G' + str(p+2),city_lst[p])
    worksheet.write('H' + str(p+2),phone_lst[p])


workbook.close()
