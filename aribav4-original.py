#!/usr/bin/python3

from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter

strasse_lst =[]
plz_lst = []
ort_lst = []
telefon_lst = []
vorname_lst = []
nachname_lst = []


workbook = xlsxwriter.Workbook('lista40.xlsx')
worksheet  = workbook.add_worksheet()


worksheet.write('A1','Anrede')
worksheet.write('B1','Titel')
worksheet.write('C1','Vorname')
worksheet.write('D1','Nachname')
worksheet.write('E1','Strasse')
worksheet.write('F1','Plz')
worksheet.write('G1','Ort')
worksheet.write('H1','Telefon')

## old https://www.herold.at/telefonbuch/villach-stadt-und-land/?page=
## ferlach == 134
## https://www.herold.at/telefonbuch/1T1bj_ferlach/?page=

##Lienz == https://www.herold.at/telefonbuch/SRKGq_lienz/?page=
#1000 - 2054
##Griffen == "https://www.herold.at/telefonbuch/1TS6Z_griffen/?page=
##https://www.herold.at/telefonbuch/jLGb3_paternion/?page=

##https://www.herold.at/telefonbuch/F57WZ_spittal-an-der-drau/?page=

##https://www.herold.at/telefonbuch/klagenfurt-land/?page=

##https://www.herold.at/telefonbuch/F6dfd_baden/?page=

##https://www.herold.at/telefonbuch/SfD9L_wiener-neustadt

##https://www.herold.at/telefonbuch/Sc56b_strasshof-an-der-nordbahn/

##https://www.herold.at/telefonbuch/DkW73_hainburg-an-der-donau/

##https://www.herold.at/telefonbuch/DsRbW_poysdorf/

##https://www.herold.at/telefonbuch/hzNFf_fischamend/

##https://www.herold.at/telefonbuch/jJKbZ_zistersdorf/

##https://www.herold.at/telefonbuch/Dnnsb_leopoldsdorf/

##https://www.herold.at/telefonbuch/DvrTS_sollenau/

##https://www.herold.at/telefonbuch/Dnnsb_leopoldsdorf/

##https://www.herold.at/telefonbuch/j5bz4_leobersdorf/

##https://www.herold.at/telefonbuch/j5bZs_leobendorf/

##https://www.herold.at/telefonbuch/SpjZQ_bisamberg/

##https://www.herold.at/telefonbuch/hzJ8q_felixdorf/

##https://www.herold.at/telefonbuch/jB6kC_retz/

##https://www.herold.at/telefonbuch/hz89c_enzesfeld-lindabrunn

##https://www.herold.at/telefonbuch/DpR7C_mannersdorf-am-leithagebirge/

##https://www.herold.at/telefonbuch/SNZmL_hinterbr%C3%BChl/

##https://www.herold.at/telefonbuch/198tm_lanzenkirchen/

##https://www.herold.at/telefonbuch/TN425_harmannsdorf/

##https://www.herold.at/telefonbuch/Slz4P_gaweinstal/

##https://www.herold.at/telefonbuch/hxbVh_ebergassing/

##https://www.herold.at/telefonbuch/DvkC6_sierndorf/

##https://www.herold.at/telefonbuch/SMfnG_gumpoldskirchen/

##https://www.herold.at/telefonbuch/jGGnN_trumau/

##https://www.herold.at/telefonbuch/SHtJg_bad-fischau-brunn

##https://www.herold.at/telefonbuch/SHS1n_angern-an-der-march/

##https://www.herold.at/telefonbuch/Dz8wg_theresienfeld/

##https://www.herold.at/telefonbuch/1C7Rl_neufeld-an-der-leitha/

##https://www.herold.at/telefonbuch/j45TS_katzelsdorf/

##https://www.herold.at/telefonbuch/SM6GH_gramatneusiedl/

##https://www.herold.at/telefonbuch/1Dr5N_pinggau/

##https://www.herold.at/telefonbuch/F7GvS_ernstbrunn/

##https://www.herold.at/telefonbuch/hxXvK_ebenfurth/

##https://www.herold.at/telefonbuch/hz6vb_enzersdorf-an-der-fischa/

##https://www.herold.at/telefonbuch/Dmffr_kittsee/

##https://www.herold.at/telefonbuch/F7G7t_bad-erlach/

##https://www.herold.at/telefonbuch/DsDnL_markt-piesting/

##https://www.herold.at/telefonbuch/SLS3H_g%C3%B6llersdorf/

##https://www.herold.at/telefonbuch/FP3Cd_bruckneudorf/

##https://www.herold.at/telefonbuch/SNzjR_hornstein/

##https://www.herold.at/telefonbuch/j6J7w_marchegg/

##https://www.herold.at/telefonbuch/1F1VC_pottenstein/

##https://www.herold.at/telefonbuch/hwb19_biedermannsdorf/

##https://www.herold.at/telefonbuch/SQ3Qm_kirchschlag-in-der-buckligen-welt/

##https://www.herold.at/telefonbuch/ScnV5_trautmannsdorf-an-der-leitha/

##https://www.herold.at/telefonbuch/j5TBb_laxenburg/

##https://www.herold.at/telefonbuch/SRDgr_leopoldsdorf-im-marchfelde/

##https://www.herold.at/telefonbuch/5PhvP_matzen-raggendorf/

##https://www.herold.at/telefonbuch/3Mgfz_wienerwald/

##https://www.herold.at/telefonbuch/jQ58T_hohenau-an-der-march/

##https://www.herold.at/telefonbuch/SR1Qq_lassee/

##https://www.herold.at/telefonbuch/Dsp6t_raabs-an-der-thaya/

##https://www.herold.at/telefonbuch/SH9Q6_alland/


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
        ort = ""
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


############    PLZ     ###########
        stradd_lst = stradd.split(' ')
        plz = ''
        for g in range(len(stradd_lst)):
            if(len(stradd_lst[g]) == 4 and stradd_lst[g].isnumeric()):
                plz = stradd_lst[g]
        if(plz == ''):
            continue
############    PLZ END     ######


############    ORT     ##########
        for q in range(stradd_lst.index(plz)+1,len(stradd_lst)):
            ort += stradd_lst[q] + " "
############    ORT End     ######


############    TEL     ##########
        if(len(tel[i].findChildren()) <= 1):
             continue
        else:
            strtel = list(tel[i].div.div.div.div)[1]
############    TEL End     ######


###### Remove spaces at the beginning and end of the str-address#########
        if(len(stradd_comma) > 2):
            strasse = stradd_comma[1][1:]
        else:
            if(stradd_comma[0][len(stradd_comma[0])-1] == ' '):
                strasse = stradd_comma[0][:-1]
            else:
                strasse = stradd_comma[0]
###### Remove spaces at the beginning and end of the str-address   END ###

##############Check if tel number starts with 6 ##########################
        strtel_lst = strtel.replace(" ","")
        telefon = strtel_lst[1:].replace('-','')
        if(telefon[0] != '6'):
            continue
##############Check if tel number starts with 6 ##########################

###### Check if client has 3 or more names  ##############################
        if(len(strname_lst) < 3):
            vorname = strname_lst[1]
            nachname = strname_lst[0]
        else:
            vorname = "**" + strname_lst[1]
            nachname = strname_lst[0]
###### Check if client has 3 or more names END ############################

        vorname_lst.append(vorname)
        nachname_lst.append(nachname)
        telefon_lst.append(int(telefon))
        strasse_lst.append(strasse)
        plz_lst.append(int(plz))
        ort_lst.append(ort[:-1])


for p in range(len(vorname_lst)):
    worksheet.write('C' + str(p+2),vorname_lst[p])
    worksheet.write('D' + str(p+2),nachname_lst[p])
    worksheet.write('E' + str(p+2),strasse_lst[p])
    worksheet.write('F' + str(p+2),plz_lst[p])
    worksheet.write('G' + str(p+2),ort_lst[p])
    worksheet.write('H' + str(p+2),telefon_lst[p])
   #print(vorname_lst[p] + " " + nachname_lst[p] + " " + str(telefon_lst[p]) + " " + strasse_lst[p] + " " + str(plz_lst[p]) + " " + ort_lst[p] + "\n\n")


workbook.close()
