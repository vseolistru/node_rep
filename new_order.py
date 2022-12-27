import datetime
import time
import xlrd
import xlwt 
from xlwt import Workbook
import csv
from datetime import timedelta
import httplib2 
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from moduls.writer_xls import Writer_4, Writer_5
from moduls.getters_csv import OrderReader, OrderTrackList_5, OrderTrackList_4, OrderTrackList_2
"""version 2.09"""

date = datetime.datetime.today().strftime("%d.%m.%y")
now = datetime.datetime.today()
one_day= timedelta(1)
next_date = now+one_day
new_date=next_date.strftime("%d.%m.%y")
#_________________________________________________________________
def download_sheet_to_cs(spreadsheet_id, sheet_name):
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range='A:AC').execute()
    output_file = 'order_list.csv'
    with open(output_file, 'w') as f:
        writer = csv.writer(f)
        writer.writerows(result.get('values'))
    f.close()
    print('Загрузил order_list.csv')
#______________________________________________________________
CREDENTIALS_FILE = 'creds.json'  
spreadsheet_id='1lH6zhhEpEoUoz3sqcUjZAkt8P_RQOgk_2oHi5Ny_Sy0'
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) 
download_sheet_to_cs(spreadsheet_id, 'Sheet1')

#_________create-class-instances____&-writing-to-xls_______
file_name_1 = 'Nakladnie-Mar-1.xls'
name1 = 'name'
point_name_1 = 'Интернет магазин'
point1 = 'internet-magazin'
Point1 = OrderReader(name1, point1).reader_quantity()
Name1 = OrderReader(name1, point1).reader_name()
#_______________________________________
name2 = 'name'
point_name_2 = 'Пушкино Чехова 40(7)'
point2 = 'pushkino-cheh-40-7'
Point2 = OrderReader(name2, point2).reader_quantity()
Name2 = OrderReader(name2, point2).reader_name()
#________________________________________
name3 = 'name'
point_name_3 = 'Пушкино Московский проспект'
point3 = 'pushkino-mosk-prosp'
Point3 = OrderReader(name3, point3).reader_quantity()
Name3 = OrderReader(name3, point3).reader_name()
#________________________________________
name4 = 'name'
point_name_4 = 'Пушкино Облака'
point4 = 'pushkino-oblaka'
Point4 = OrderReader(name4, point4).reader_quantity()
Name4 = OrderReader(name4, point4).reader_name()
#________________________________________
point_name_5 = 'Пушкино Просвещения 13 к.3'
name5 = 'name'
point5 = 'pushkino-prosveshenia'
Point5 = OrderReader(name5, point5).reader_quantity()
Name5 = OrderReader(name5, point5).reader_name()
#________________________________________writng_____________
Nakladnie_1 = Writer_5(point_name_1, point_name_2, point_name_3, point_name_4, point_name_5, Name1, Point1, Name2, Point2, Name3, Point3, Name4, Point4, Name5, Point5, file_name_1).write_goods_list_5() 
#___________________________________________________________
file_name_2 = 'Nakladnie-Mar-2.xls'
point_name_6 = 'С-Посад Ферма'
name6 = 'name'
point6 = 'sposad-ferma'
Point6 = OrderReader(name6, point6).reader_quantity()
Name6 = OrderReader(name6, point6).reader_name()
#________________________________________
point_name_7 = 'С-Посад Скобянка'
name7 = 'name'
point7 = 'sposad-skobyanka'
Point7 = OrderReader(name7, point7).reader_quantity()
Name7 = OrderReader(name7, point7).reader_name()
#________________________________________
point_name_8 = 'С-Посад Кр-армии-182'
name8 = 'name'
point8 = 's-posad-armii-182'
Point8 = OrderReader(name8, point8).reader_quantity()
Name8 = OrderReader(name8, point8).reader_name()
#________________________________________
point_name_9 = 'С-Посад Кр-армии-203'
name9 = 'name'
point9 = 'sposad-armii-203'
Point9 = OrderReader(name9, point9).reader_quantity()
Name9 = OrderReader(name9, point9).reader_name()
#_______________________________________________________________
Nakladnie_2 = Writer_4(point_name_6, point_name_7, point_name_8, point_name_9, Name6, Point6, Name7, Point7, Name8, Point8, Name9, Point9, file_name_2).write_goods_list_4() 
#_______________________________________________________________
file_name_3 = 'Nakladnie-Mar-3.xls'
point_name_11 = 'Королев Горького - 6'
name11 = 'name'
point11 = 'korolev-gor-6'
Point11 = OrderReader(name11, point11).reader_quantity()
Name11 = OrderReader(name11, point11).reader_name()
#________________________________________
point_name_12 = 'Черноголовка'
name12 = 'name'
point12 = 'chernogolovka'
Point12 = OrderReader(name12, point12).reader_quantity()
Name12 = OrderReader(name12, point12).reader_name()
#_________________________________________
point_name_13 = 'Ивантеевка Бережок'
name13 = 'name'
point13 = 'ivanteevka-beregok'
Point13 = OrderReader(name13, point13).reader_quantity()
Name13 = OrderReader(name13, point13).reader_name()
#_________________________________________
point_name_14 = 'Черноголовка2'
name14 = 'name'
point14 = 'chernogolovka2'
Point14 = OrderReader(name14, point14).reader_quantity()
Name14 = OrderReader(name14, point14).reader_name()
#___________________________________________________________________
Nakladnie_3 = Writer_4(point_name_11, point_name_12, point_name_13, point_name_14, Name11, Point11, Name12, Point12, Name13, Point13, Name14, Point14, file_name_3).write_goods_list_4()
#___________________________________________________________________
file_name_4 = 'Nakladnie-Mar-4.xls'
point_name_15 = 'Софрино'
name15 = 'name'
point15 = 'sofrino'
Point15 = OrderReader(name15, point15).reader_quantity()
Name15 = OrderReader(name15, point15).reader_name()
#________________________________________
point_name_16 = 'С-Посад Кр-армии-185'
name16 = 'name'
point16 = 's-posad-armii-185'
Point16 = OrderReader(name16, point16).reader_quantity()
Name16 = OrderReader(name16, point16).reader_name()
#________________________________________
point_name_17 = 'С-Посад Кр-армии-5'
name17 = 'name'
point17 = 'sposad-armii-5'
Point17 = OrderReader(name17, point17).reader_quantity()
Name17 = OrderReader(name17, point17).reader_name()
#________________________________________
point_name_18 = 'С-Посад Кр-армии-3'
name18 = 'name'
point18 = 'sposad-armii-3'
Point18 = OrderReader(name18, point18).reader_quantity()
Name18 = OrderReader(name18, point18).reader_name()
#______________________________________________________________
Nakladnie_4 = Writer_4(point_name_15, point_name_16, point_name_17, point_name_18, Name15, Point15, Name16, Point16, Name17, Point17, Name18, Point18, file_name_4).write_goods_list_4()
#______________________________________________________________
file_name_5 = 'Nakladnie-Mar-5.xls'
point_name_19 = 'П.Набережная'
name19 = 'name'
point19 = 'pnaberezhnay'
Point19 = OrderReader(name19, point19).reader_quantity()
Name19 = OrderReader(name19, point19).reader_name()
#________________________________________
point_name_20 = 'Заветы Ильича'
name20 = 'name'
point20 = 'zavety-ilicha'
Point20 = OrderReader(name20, point20).reader_quantity()
Name20 = OrderReader(name20, point20).reader_name()
#________________________________________
point_name_21 = 'Пушкино Арманд'
name21 = 'name'
point21 = 'pushkino-armand'
Point21 = OrderReader(name21, point21).reader_quantity()
Name21 = OrderReader(name21, point21).reader_name()
#________________________________________
point_name_22 = 'Пушкино Тургенева 18'
name22 = 'name'
point22 = 'pushkino-turgeneva'
Point22 = OrderReader(name22, point22).reader_quantity()
Name22 = OrderReader(name22, point22).reader_name()
#________________________________________________________________
Nakladnie_4 = Writer_4(point_name_19, point_name_20, point_name_21, point_name_22, Name19, Point19, Name20, Point20, Name21, Point21, Name22, Point22, file_name_5).write_goods_list_4()
#________________ending-writing__________________________________________

#____________________create-class-instances-for-tracklist_____________
trackname1 = 'name'
trackpointa_1 = 'internet-magazin'
trackpointb_1 = 'pushkino-cheh-40-7'
trackpointc_1 = 'pushkino-mosk-prosp'
trackpointd_1 = 'pushkino-oblaka'
trackpointe_1 = 'pushkino-prosveshenia'
trackpointf_1 = 'vsego-1'
TrackName1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_name()
TrackPointa_1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_1()
TrackPointb_1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_2()
TrackPointc_1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_3()
TrackPointd_1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_4()
TrackPointe_1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_5()
TrackPointf_1 = OrderTrackList_5(trackname1, trackpointa_1, trackpointb_1, trackpointc_1, trackpointd_1, trackpointe_1, trackpointf_1).reader_qua_ty_6()
#_________________________________________________________________
trackname2 = 'name'
trackpointa_2 = 'sposad-ferma'
trackpointb_2 = 'sposad-skobyanka'         
trackpointc_2 = 's-posad-armii-182'
trackpointd_2 = 'sposad-armii-203'
trackpointf_2 = 'vsego-2'
TrackName2 = OrderTrackList_4(trackname2, trackpointa_2, trackpointb_2, trackpointc_2, trackpointd_2, trackpointf_2).reader_qua_ty_name()
TrackPointa_2 = OrderTrackList_4(trackname2, trackpointa_2, trackpointb_2, trackpointc_2, trackpointd_2, trackpointf_2).reader_qua_ty_1()
TrackPointb_2 = OrderTrackList_4(trackname2, trackpointa_2, trackpointb_2, trackpointc_2, trackpointd_2, trackpointf_2).reader_qua_ty_2()
TrackPointc_2 = OrderTrackList_4(trackname2, trackpointa_2, trackpointb_2, trackpointc_2, trackpointd_2, trackpointf_2).reader_qua_ty_3()
TrackPointd_2 = OrderTrackList_4(trackname2, trackpointa_2, trackpointb_2, trackpointc_2, trackpointd_2, trackpointf_2).reader_qua_ty_4()
TrackPointf_2 = OrderTrackList_4(trackname2, trackpointa_2, trackpointb_2, trackpointc_2, trackpointd_2, trackpointf_2).reader_qua_ty_5()
#___________________________________________________________________
trackname3 = 'name'
trackpointa_3 = 'korolev-gor-6'
trackpointb_3 = 'chernogolovka'
trackpointc_3 = 'ivanteevka-beregok'
trackpointd_3 = 'chernogolovka2'
trackpointe_3 = 'vsego-3'
TrackName3 = OrderTrackList_4(trackname3, trackpointa_3, trackpointb_3, trackpointc_3, trackpointd_3, trackpointe_3).reader_qua_ty_name()
TrackPointa_3 = OrderTrackList_4(trackname3, trackpointa_3, trackpointb_3, trackpointc_3, trackpointd_3, trackpointe_3).reader_qua_ty_1()
TrackPointb_3 = OrderTrackList_4(trackname3, trackpointa_3, trackpointb_3, trackpointc_3, trackpointd_3, trackpointe_3).reader_qua_ty_2()
TrackPointc_3 = OrderTrackList_4(trackname3, trackpointa_3, trackpointb_3, trackpointc_3, trackpointd_3, trackpointe_3).reader_qua_ty_3()
TrackPointd_3 = OrderTrackList_4(trackname3, trackpointa_3, trackpointb_3, trackpointc_3, trackpointd_3, trackpointe_3).reader_qua_ty_4()
TrackPointe_3 = OrderTrackList_4(trackname3, trackpointa_3, trackpointb_3, trackpointc_3, trackpointd_3, trackpointe_3).reader_qua_ty_5()
#_________________________________________________________________________
trackname4 = 'name'
trackpointa_4 = 'sofrino'
trackpointb_4 = 's-posad-armii-185'
trackpointc_4 = 'sposad-armii-5'
trackpointd_4 = 'sposad-armii-3'
trackpointe_4 = 'vsego-4'
TrackName4 = OrderTrackList_4(trackname4, trackpointa_4, trackpointb_4, trackpointc_4, trackpointd_4, trackpointe_4).reader_qua_ty_name()
TrackPointa_4 = OrderTrackList_4(trackname4, trackpointa_4, trackpointb_4, trackpointc_4, trackpointd_4, trackpointe_4).reader_qua_ty_1()
TrackPointb_4 = OrderTrackList_4(trackname4, trackpointa_4, trackpointb_4, trackpointc_4, trackpointd_4, trackpointe_4).reader_qua_ty_2()
TrackPointc_4 = OrderTrackList_4(trackname4, trackpointa_4, trackpointb_4, trackpointc_4, trackpointd_4, trackpointe_4).reader_qua_ty_3()
TrackPointd_4 = OrderTrackList_4(trackname4, trackpointa_4, trackpointb_4, trackpointc_4, trackpointd_4, trackpointe_4).reader_qua_ty_4()
TrackPointe_4 = OrderTrackList_4(trackname4, trackpointa_4, trackpointb_4, trackpointc_4, trackpointd_4, trackpointe_4).reader_qua_ty_5()
#____________________________________________________________________________
trackname5 = 'name'
trackpointa_5 = 'zavety-ilicha'
trackpointb_5 = 'pushkino-armand'
trackpointc_5 = 'pushkino-turgeneva'
trackpointd_5 = 'pnaberezhnay'
trackpointe_5 = 'vsego-5'
TrackName5 = OrderTrackList_4(trackname5, trackpointa_5, trackpointb_5, trackpointc_5, trackpointd_5, trackpointe_5).reader_qua_ty_name()
TrackPointa_5 = OrderTrackList_4(trackname5, trackpointa_5, trackpointb_5, trackpointc_5, trackpointd_5, trackpointe_5).reader_qua_ty_1()
TrackPointb_5 = OrderTrackList_4(trackname5, trackpointa_5, trackpointb_5, trackpointc_5, trackpointd_5, trackpointe_5).reader_qua_ty_2()
TrackPointc_5 = OrderTrackList_4(trackname5, trackpointa_5, trackpointb_5, trackpointc_5, trackpointd_5, trackpointe_5).reader_qua_ty_3()
TrackPointd_5 = OrderTrackList_4(trackname5, trackpointa_5, trackpointb_5, trackpointc_5, trackpointd_5, trackpointe_5).reader_qua_ty_4()
TrackPointe_5 = OrderTrackList_4(trackname5, trackpointa_5, trackpointb_5, trackpointc_5, trackpointd_5, trackpointe_5).reader_qua_ty_5()
#________________________________________________________________________________
trackname6 = 'name'
trackpointa_6 = 'id'
trackpointb_6 = 'sklad'
TrackName6 = OrderTrackList_2(trackname6, trackpointa_6, trackpointb_6).reader_qua_ty_name()
TrackPointa_6 = OrderTrackList_2(trackname6, trackpointa_6, trackpointb_6).reader_qua_ty_1()
TrackPointb_6 = OrderTrackList_2(trackname6, trackpointa_6, trackpointb_6).reader_qua_ty_2()
#_________________________________________________________________________________________________________
#_________________________________________write-tacklist__________________________________________________
wb=Workbook()
style_string='font:bold on; borders: bottom dashed; borders: top dotted; borders: right dashed;'
style_string_1='font:bold on; borders: bottom dashed; borders: top dotted;'
style1=xlwt.easyxf(style_string_1)
style=xlwt.easyxf(style_string)
#_______________________________________________________________
sheet2 = wb.add_sheet('Маршрут -1')
sheet2.write(0, 2, 'Маршрут - 1')
sheet2.write(0, 4, 'Дата:', style=style1)
sheet2.write(0, 5, new_date, style=style)
sheet2.write(2, 0, 'Название точек:')
sheet2.write(2, 1, 'Пушкино: Инт-магаз, Чехова 40/7, Моск-проспект, Облака, Просвещения', style=style1)
sheet2.write(4, 0, 'позиции для отгрузки', style=style1)
sheet2.write(4, 1, '', style=style1)
sheet2.write(4, 2, 'Инт-магаз', style=style)
sheet2.write(4, 3, 'Пуш-Чех 40/7', style=style)
sheet2.write(4, 4, 'Пуш-Моск пр', style=style)
sheet2.write(4, 5, 'Пуш-Облак', style=style)
sheet2.write(4, 6, 'Пуш-Просв', style=style)
sheet2.write(4, 7, 'Всего:', style=style)
if len(TrackName1)>5:
   a=TrackName1[0:55]
   b=TrackName1[55:114]
   c=TrackName1[114:173]
   d=TrackName1[173:232]
   e=TrackName1[232:291]
   row=5
   col=0
   for x in a:   
     sheet2.write(row, 0, x, style=style1)
     sheet2.write(row, 1, '', style=style1)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet2.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(60, 1, '', style=style1)
     for x in b:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet2.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(120, 1, '', style=style1)
     for x in c:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet2.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(180, 1, '', style=style1)
     for x in d:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet2.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet2.write(240, 1, '', style=style1)
     for x in e:
       sheet2.write(row, 0, x, style=style1)
       sheet2.write(row, 1, '', style=style1)
       row+=1
if len(TrackPointa_1)>5:
   a=TrackPointa_1[0:55]
   b=TrackPointa_1[55:114]
   c=TrackPointa_1[114:173]
   d=TrackPointa_1[173:232]
   e=TrackPointa_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 2, 'Инт-магаз', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 2, 'Инт-магаз', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 2, 'Инт-магаз', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 2, 'Инт-магаз', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 2, y, style=style)
       row+=1 
if len(TrackPointb_1)>5:
   a=TrackPointb_1[0:55]
   b=TrackPointb_1[55:114]
   c=TrackPointb_1[114:173]
   d=TrackPointb_1[173:232]
   e=TrackPointb_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 3, 'Пуш-Чех 40/7', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 3, 'Пуш-Чех 40/7', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 3, 'Пуш-Чех 40/7', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 3, 'Пуш-Чех 40/7', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 3, y, style=style)
       row+=1 
if len(TrackPointc_1)>5:
   a=TrackPointc_1[0:55]
   b=TrackPointc_1[55:114]
   c=TrackPointc_1[114:173]
   d=TrackPointc_1[173:232]
   e=TrackPointc_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 4, 'Пуш-Моск пр', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 4, 'Пуш-Моск пр', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 4, 'Пуш-Моск пр', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 4, 'Пуш-Моск пр', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 4, y, style=style)
       row+=1 
if len(TrackPointd_1)>5:
   a=TrackPointd_1[0:55]
   b=TrackPointd_1[55:114]
   c=TrackPointd_1[114:173]
   d=TrackPointd_1[173:232]
   e=TrackPointd_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 5, 'Пуш-Облак', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 5, 'Пуш-Облак', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 5, 'Пуш-Облак', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 5, 'Пуш-Облак', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 5, y, style=style)
       row+=1 
if len(TrackPointe_1)>5:
   a=TrackPointe_1[0:55]
   b=TrackPointe_1[55:114]
   c=TrackPointe_1[114:173]
   d=TrackPointe_1[173:232]
   e=TrackPointe_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 6, 'Пуш-Просв', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 6, 'Пуш-Просв', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 6, 'Пуш-Просв', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 6, 'Пуш-Просв', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 6, y, style=style)
       row+=1 
if len(TrackPointf_1)>5:
   a=TrackPointf_1[0:55]
   b=TrackPointf_1[55:114]
   c=TrackPointf_1[114:173]
   d=TrackPointf_1[173:232]
   e=TrackPointf_1[232:291]  
   row=5
   col=0
   for y in a:
    sheet2.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet2.write(60, 7, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet2.write(120, 7, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet2.write(180, 7, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet2.write(240, 7, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet2.write(row, 7, y, style=style)
       row+=1 
   row+=2
   sheet2.write(row, 0, 'Получил:')
   sheet2.write(row, 2, '________________/_______________/')
   row+=1
#__________________________________________
sheet3 = wb.add_sheet('Маршрут -2') 
sheet3.write(0, 2, 'Маршрут - 2')
sheet3.write(0, 4, 'Дата:', style=style1)
sheet3.write(0, 5, new_date, style=style)
sheet3.write(2, 0, 'Название точек:')
sheet3.write(2, 1, 'СПосад Ферма, С-Посад Скобянка, Кр-армии-182, Кр-армии-203', style=style1)
sheet3.write(4, 0, 'позиции для отгрузки', style=style1)
sheet3.write(4, 1, '', style=style1)
sheet3.write(4, 2, 'СПос Ферма', style=style)
sheet3.write(4, 3, 'СПос Скоб', style=style)
sheet3.write(4, 4, 'СП Ар-182', style=style)
sheet3.write(4, 5, 'СП Ар-203', style=style)
sheet3.write(4, 6, 'Всего:', style=style)
if len(TrackName2)>5:
   a=TrackName2[0:55]
   b=TrackName2[55:114]
   c=TrackName2[114:173]
   d=TrackName2[173:232]
   e=TrackName2[232:291]
   row=5
   col=0
   for x in a:   
     sheet3.write(row, 0, x, style=style1)
     sheet3.write(row, 1, '', style=style)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet3.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(60, 1, '', style=style)
     for x in b:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet3.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(120, 1, '', style=style)
     for x in c:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet3.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(180, 1, '', style=style)
     for x in d:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet3.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet3.write(240, 1, '', style=style)
     for x in e:
       sheet3.write(row, 0, x, style=style1)
       sheet3.write(row, 1, '', style=style)
       row+=1
if len(TrackPointa_2)>5:
   a=TrackPointa_2[0:55]
   b=TrackPointa_2[55:114]
   c=TrackPointa_2[114:173]
   d=TrackPointa_2[173:232]
   e=TrackPointa_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 2, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60,2, 'СПос Ферма', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 2, 'СПос Ферма', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 2, 'СПос Ферма', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 2, 'СПос Ферма', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 2, y, style=style)
       row+=1 
if len(TrackPointb_2)>5:
   a=TrackPointb_2[0:55]
   b=TrackPointb_2[55:114]
   c=TrackPointb_2[114:173]
   d=TrackPointb_2[173:232]
   e=TrackPointb_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 3, 'СП Скоб', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 3, 'СП Скоб', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 3, 'СП Скоб', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 3, 'СП Скоб', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 3, y, style=style)
       row+=1
if len(TrackPointc_2)>5:
   a=TrackPointc_2[0:55]
   b=TrackPointc_2[55:114]
   c=TrackPointc_2[114:173]
   d=TrackPointc_2[173:232]
   e=TrackPointc_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 4, 'СП Ар-182', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 4, 'СП Ар-182', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 4, 'СП Ар-182', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 4, 'СП Ар-182', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 4, y, style=style)
       row+=1 
if len(TrackPointd_2)>5:
   a=TrackPointd_2[0:55]
   b=TrackPointd_2[55:114]
   c=TrackPointd_2[114:173]
   d=TrackPointd_2[173:232]
   e=TrackPointd_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 5, 'СП Ар-203', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 5, 'СП Ар-203', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 5, 'СП Ар-203', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 5, 'СП Ар-203', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 5, y, style=style)
       row+=1 

if len(TrackPointf_2)>5:
   a=TrackPointf_2[0:55]
   b=TrackPointf_2[55:114]
   c=TrackPointf_2[114:173]
   d=TrackPointf_2[173:232]
   e=TrackPointf_2[232:291]  
   row=5
   col=0
   for y in a:
    sheet3.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet3.write(60, 6, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet3.write(120, 6, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet3.write(180, 6, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet3.write(240, 6, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet3.write(row, 6, y, style=style)
       row+=1 
   row+=2
   sheet3.write(row, 0, 'Получил:')
   sheet3.write(row, 2, '________________/_______________/')
   row+=1
#____________________________________
sheet4 = wb.add_sheet('Маршрут -3') 
sheet4.write(0, 2, 'Маршрут - 3')
sheet4.write(0, 4, 'Дата:', style=style1)
sheet4.write(0, 5, new_date, style=style)
sheet4.write(2, 0, 'Название точек:')
sheet4.write(2, 1, 'Королев, Ивантеевка Бережок, Черноголовка, Черноголовка2', style=style1)
sheet4.write(4, 1, '', style=style1)
sheet4.write(4, 2, '', style=style)
sheet4.write(4, 0, 'позиции для отгрузки', style=style1)
sheet4.write(4, 3, 'Королев', style=style)
sheet4.write(4, 4, 'Ив Бережок', style=style)
sheet4.write(4, 5, 'Черноголов', style=style)
sheet4.write(4, 6, 'Черноголов2', style=style)
sheet4.write(4, 7, 'Всего:', style=style)
if len(TrackName3)>5:
   a=TrackName3[0:55]
   b=TrackName3[55:114]
   c=TrackName3[114:173]
   d=TrackName3[173:232]
   e=TrackName3[232:291]
   row=5
   col=0
   for x in a:   
     sheet4.write(row, 0, x, style=style1)
     sheet4.write(row, 1, '', style=style1)
     sheet4.write(row, 2, '', style=style)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet4.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(60, 1, '', style=style1)
     sheet4.write(60, 2, '', style=style)
     for x in b:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)
       sheet4.write(row, 2, '', style=style)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet4.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(120, 1, '', style=style1)
     sheet4.write(120, 2, '', style=style)
     for x in c:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)
       sheet4.write(row, 2, '', style=style)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet4.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(180, 1, '', style=style1)
     sheet4.write(180, 2, '', style=style)
     for x in d:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)
       sheet4.write(row, 2, '', style=style)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet4.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet4.write(240, 1, '', style=style1)
     sheet4.write(240, 2, '', style=style)
     for x in e:
       sheet4.write(row, 0, x, style=style1)
       sheet4.write(row, 1, '', style=style1)
       sheet4.write(row, 2, '', style=style)
       row+=1
if len(TrackPointa_3)>5:#____________Королев
   a=TrackPointa_3[0:55]
   b=TrackPointa_3[55:114]
   c=TrackPointa_3[114:173]
   d=TrackPointa_3[173:232]
   e=TrackPointa_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 3, 'Королев', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 3, 'Королев', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 3, 'Королев', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 3, 'Королев', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 3, y, style=style)
       row+=1
if len(TrackPointc_3)>5:
   a=TrackPointc_3[0:55]
   b=TrackPointc_3[55:114]
   c=TrackPointc_3[114:173]
   d=TrackPointc_3[173:232]
   e=TrackPointc_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 4, 'Ив Бережок', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 4, 'Ив Бережок', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 4, 'Ив Бережок', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 4, 'Ив Бережок', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 4, y, style=style)
       row+=1 
if len(TrackPointb_3)>5:
   a=TrackPointb_3[0:55]
   b=TrackPointb_3[55:114]
   c=TrackPointb_3[114:173]
   d=TrackPointb_3[173:232]
   e=TrackPointb_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 5, 'Черноголов', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 5, 'Черноголов', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 5, 'Черноголов', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 5, 'Черноголов', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 5, y, style=style)
       row+=1

if len(TrackPointd_3)>5: 
   a=TrackPointd_3[0:55]
   b=TrackPointd_3[55:114]
   c=TrackPointd_3[114:173]
   d=TrackPointd_3[173:232]
   e=TrackPointd_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 6, 'Черноголов2', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 6, 'Черноголов2', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 6, 'Черноголов2', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 6, 'Черноголов2', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 6, y, style=style)
       row+=1 

if len(TrackPointe_3)>5:
   a=TrackPointe_3[0:55]
   b=TrackPointe_3[55:114]
   c=TrackPointe_3[114:173]
   d=TrackPointe_3[173:232]
   e=TrackPointe_3[232:291]  
   row=5
   col=0
   for y in a:
    sheet4.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet4.write(60, 7, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet4.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet4.write(120, 7, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet4.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet4.write(180, 7, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet4.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet4.write(240, 7, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet4.write(row, 7, y, style=style)
       row+=1
   row+=2
   sheet4.write(row, 0, 'Получил:')
   sheet4.write(row, 2, '________________/_______________/')
   row+=1
#____________________________________
sheet5 = wb.add_sheet('Маршрут -4') 
sheet5.write(0, 2, 'Маршрут - 4')
sheet5.write(0, 4, 'Дата:', style=style1)
sheet5.write(0, 5, new_date, style=style)
sheet5.write(2, 0, 'Название точек:')
sheet5.write(2, 1, 'Софрино, С-Посад Кр-армии-185, С-Посад Кр-армии-5,С-Посад Кр-армии-3,', style=style1)
sheet5.write(4, 0, 'позиции для отгрузки', style=style1)
sheet5.write(4, 1, '', style=style1)
sheet5.write(4, 2, '', style=style)
sheet5.write(4, 3, 'Софрино', style=style)
sheet5.write(4, 4, 'СП Ар-185', style=style)
sheet5.write(4, 5, 'СП Кр.ар-5', style=style)
sheet5.write(4, 6, 'СП Кр.ар-3', style=style)
sheet5.write(4, 7, 'Всего:', style=style)
if len(TrackName4)>5:
   a=TrackName4[0:55]
   b=TrackName4[55:114]
   c=TrackName4[114:173]
   d=TrackName4[173:232]
   e=TrackName4[232:291]
   row=5
   col=0
   for x in a:   
     sheet5.write(row, 0, x, style=style1)
     sheet5.write(row, 1, '', style=style1)
     sheet5.write(row, 2, '', style=style)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet5.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(60, 1, '', style=style1)
     sheet5.write(60, 2, '', style=style)
     for x in b:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       sheet5.write(row, 2, '', style=style)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet5.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(120, 1, '', style=style1)
     sheet5.write(120, 2, '', style=style)
     for x in c:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       sheet5.write(row, 2, '', style=style)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet5.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(180, 1, '', style=style1)
     sheet5.write(180, 2, '', style=style)
     for x in d:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       sheet5.write(row, 2, '', style=style)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet5.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet5.write(240, 1, '', style=style1)
     sheet5.write(240, 2, '', style=style)
     for x in e:
       sheet5.write(row, 0, x, style=style1)
       sheet5.write(row, 1, '', style=style1)
       sheet5.write(row, 2, '', style=style)
       row+=1
if len(TrackPointa_4)>5:
   a=TrackPointa_4[0:55]
   b=TrackPointa_4[55:114]
   c=TrackPointa_4[114:173]
   d=TrackPointa_4[173:232]
   e=TrackPointa_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 3, 'Софрино', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 3, 'Софрино', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 3, 'Софрино', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 3, 'Софрино', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 3, y, style=style)
       row+=1 
if len(TrackPointb_4)>5:
   a=TrackPointb_4[0:55]
   b=TrackPointb_4[55:114]
   c=TrackPointb_4[114:173]
   d=TrackPointb_4[173:232]
   e=TrackPointb_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 4, 'СП Кр.ар-185', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 4, 'СП Кр.ар-185', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 4, 'СП Кр.ар-185', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 4, 'СП Кр.ар-185', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 4, y, style=style)
       row+=1 
if len(TrackPointc_4)>5:
   a=TrackPointc_4[0:55]
   b=TrackPointc_4[55:114]
   c=TrackPointc_4[114:173]
   d=TrackPointc_4[173:232]
   e=TrackPointc_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 5, 'СП Кр.ар-5', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 5, 'СП Кр.ар-5', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 5, 'СП Кр.ар-5', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 5, 'СП Кр.ар-5', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 5, y, style=style)
       row+=1 
if len(TrackPointd_4)>5:
   a=TrackPointd_4[0:55]
   b=TrackPointd_4[55:114]
   c=TrackPointd_4[114:173]
   d=TrackPointd_4[173:232]
   e=TrackPointd_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 6, 'СП Кр.ар-3', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 6, 'СП Кр.ар-3', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 6, 'СП Кр.ар-3', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 6, 'СП Кр.ар-3', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 6, y, style=style)
       row+=1 
if len(TrackPointe_4)>5:
   a=TrackPointe_4[0:55]
   b=TrackPointe_4[55:114]
   c=TrackPointe_4[114:173]
   d=TrackPointe_4[173:232]
   e=TrackPointe_4[232:291]  
   row=5
   col=0
   for y in a:
    sheet5.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet5.write(60, 7, 'Всего:', style=style)
     row=61
     col=0 
     for y in b:  
       sheet5.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet5.write(120, 7, 'Всего:', style=style)
     row=121
     col=0 
     for y in c:  
       sheet5.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet5.write(180, 7, 'Всего:', style=style)
     row=181
     col=0 
     for y in d:  
       sheet5.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet5.write(240, 7, 'Всего:', style=style)
     row=241
     col=0 
     for y in e:  
       sheet5.write(row, 7, y, style=style)
       row+=1
   row+=2
   sheet5.write(row, 0, 'Получил:')
   sheet5.write(row, 2, '________________/_______________/')
   row+=1
#__________________________________________
sheet6 = wb.add_sheet('Маршрут -5', cell_overwrite_ok=True) 
sheet6.write(0, 2, 'Маршрут - 5')
sheet6.write(0, 4, 'Дата:', style=style1)
sheet6.write(0, 5, new_date, style=style)
sheet6.write(2, 0, 'Название точек:')
sheet6.write(2, 1, 'Арманд, Набережная, Тургенева -18 Заветы Ильича', style=style1)
sheet6.write(4, 0, 'позиции для отгрузки', style=style1)
sheet6.write(4, 1, '', style=style1)
sheet6.write(4, 2, '', style=style)
sheet6.write(4, 3, 'Арманд', style=style)
sheet6.write(4, 4, 'Набережная', style=style)
sheet6.write(4, 5, 'Тургенева', style=style)
sheet6.write(4, 6, 'Заветы Ил.', style=style)
sheet6.write(4, 7, 'Всего:', style=style)
if len(TrackName5)>5:
   a=TrackName5[0:55]
   b=TrackName5[55:114]
   c=TrackName5[114:173]
   d=TrackName5[173:232]
   e=TrackName5[232:291]
   row=5
   col=0
   for x in a:   
     sheet6.write(row, 0, x, style=style1)
     sheet6.write(row, 1, '', style=style1)
     sheet6.write(row, 2, '', style=style)
     row+=1
   row=61
   col=0
   if len(b)!=0:
     sheet6.write(60, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(60, 1, '', style=style1)
     sheet6.write(60, 2, '', style=style)
     for x in b:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       sheet6.write(row, 2, '', style=style)
       row+=1
   row=121
   col=0
   if len(c)!=0:
     sheet6.write(120, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(120, 1, '', style=style1)
     sheet6.write(120, 2, '', style=style)
     for x in c:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       sheet6.write(row, 2, '', style=style)
       row+=1
   row=181
   col=0
   if len(d)!=0:
     sheet6.write(180, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(180, 1, '', style=style1)
     sheet6.write(180, 2, '', style=style)
     for x in d:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       sheet6.write(row, 2, '', style=style)
       row+=1
   row=241
   col=0
   if len(e)!=0:
     sheet6.write(240, 0, 'позиции для отгрузки', style=style1)
     sheet6.write(240, 1, '', style=style1)
     sheet6.write(240, 2, '', style=style)
     for x in e:
       sheet6.write(row, 0, x, style=style1)
       sheet6.write(row, 1, '', style=style1)
       sheet6.write(row, 2, '', style=style)
       row+=1




if len(TrackPointb_5)>5:
   a=TrackPointb_5[0:55]
   b=TrackPointb_5[55:114]
   c=TrackPointb_5[114:173]
   d=TrackPointb_5[173:232]
   e=TrackPointb_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 3, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 3, 'Пуш Арманд', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 3, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 3, 'Пуш Арманд', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 3, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 3, 'Пуш Арманд', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 3, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 3, 'Пуш Арманд', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 3, y, style=style)
       row+=1 

if len(TrackPointd_5)>5:#
   a=TrackPointd_5[0:55]
   b=TrackPointd_5[55:114]
   c=TrackPointd_5[114:173]
   d=TrackPointd_5[173:232]
   e=TrackPointd_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 4, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 4, 'Набережная', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 4, 'Набережная', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 4, 'Набережная', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 4, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 4, 'Набережная', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 4, y, style=style)
       row+=1 

if len(TrackPointc_5)>5:
   a=TrackPointc_5[0:55]
   b=TrackPointc_5[55:114]
   c=TrackPointc_5[114:173]
   d=TrackPointc_5[173:232]
   e=TrackPointc_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 5, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 5, 'Пуш Тур-18', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 5, 'Пуш Тур-18', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 5, 'Пуш Тур-18', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 5, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 5, 'Пуш Тур-18', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 5, y, style=style)
       row+=1 

if len(TrackPointa_5)>5:
   a=TrackPointa_5[0:55]
   b=TrackPointa_5[55:114]
   c=TrackPointa_5[114:173]
   d=TrackPointa_5[173:232]
   e=TrackPointa_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 6, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 6, 'Заветы Ил', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 6, 'Заветы Ил', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 6, 'Заветы Ил', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 6, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 6, 'Заветы Ил', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 6, y, style=style)
       row+=1 


if len(TrackPointe_5)>5:
   a=TrackPointe_5[0:55]
   b=TrackPointe_5[55:114]
   c=TrackPointe_5[114:173]
   d=TrackPointe_5[173:232]
   e=TrackPointe_5[232:291]  
   row=5
   col=0
   for y in a:
    sheet6.write(row, 7, y, style=style)
    row+=1
   if len(b)!=0:   
     sheet6.write(60, 7, 'Всего', style=style)
     row=61
     col=0 
     for y in b:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   if len(c)!=0:   
     sheet6.write(120, 7, 'Всего', style=style)
     row=121
     col=0 
     for y in c:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   if len(d)!=0:   
     sheet6.write(180, 7, 'Всего', style=style)
     row=181
     col=0 
     for y in d:  
       sheet6.write(row, 7, y, style=style)
       row+=1 
   if len(e)!=0:   
     sheet6.write(240, 7, 'Всего', style=style)
     row=241
     col=0 
     for y in e:  
       sheet6.write(row, 7, y, style=style)
       row+=1 




   row+=2
   sheet6.write(row, 0, 'Получил:')
   sheet6.write(row, 2, '________________/_______________/')
   row+=1

#_________________________________________
sheet8 = wb.add_sheet('Общий заказ') 
sheet8.write(0, 1, 'Общий заказ')
sheet8.write(0, 3, 'Дата:', style=style1)
sheet8.write(0, 4, new_date, style=style)
sheet8.write(2, 0, 'Всего во всех заказах:', style=style1)
sheet8.write(4, 0, 'Наименование товара', style=style1)
sheet8.write(4, 1, '', style=style1)
sheet8.write(4, 2, 'Ед.Изм', style=style)
sheet8.write(4, 3, 'Всего:', style=style)
if len(TrackName6)>5:
   a=TrackName6[0:55]
   b=TrackName6[55:114]
   c=TrackName6[114:173]
   d=TrackName6[173:232]
   e=TrackName6[232:291]
   row=5
   col=1
   for x in a:  
     sheet8.write(row, 0, x, style=style1)
     row+=1  
   row=61
   col=1
   if len(b)!=0:
       sheet8.write(60, 0, 'Наименование товара', style=style1)
       for x in b:  
         sheet8.write(row, 0, x, style=style1)
         row+=1    
   row=121
   col=1
   if len(c)!=0:
      sheet8.write(120, 0, 'Наименование товара', style=style1)
      for x in c:  
        sheet8.write(row, 0, x, style=style1)
        row+=1    
   row=181
   col=1   
   if len(d)!=0:
      sheet8.write(180, 0, 'Наименование товара', style=style1)
      for x in d:  
        sheet8.write(row, 0, x, style=style1)
        row+=1 
   row=241
   col=1
   if len(e)!=0:
      sheet8.write(240, 0, 'Наименование товара', style=style1)
      for x in e:  
        sheet8.write(row, 0, x, style=style1)
        row+=1 
if len(TrackPointa_6)>5:
   a=TrackPointa_6[0:55]
   b=TrackPointa_6[55:114]
   c=TrackPointa_6[114:173]
   d=TrackPointa_6[173:232]
   e=TrackPointa_6[232:291]
   row=5
   col=1
   for z in a:  
     sheet8.write(row, 2, z, style=style1)
     row+=1  
   row=61
   col=1
   if len(b)!=0:
       sheet8.write(60, 2, 'Ед.Изм', style=style1)
       for z in b:  
         sheet8.write(row, 2, z, style=style1)
         row+=1    
   row=121
   col=1
   if len(c)!=0:
      sheet8.write(120, 2, 'Ед.Изм', style=style1)
      for z in c:  
        sheet8.write(row, 2, z, style=style1)
        row+=1 
   
   row=181
   col=1   
   if len(d)!=0:
      sheet8.write(180, 2, 'Ед.Изм', style=style1)
      for z in d:  
        sheet8.write(row, 2, z, style=style1)
        row+=1 
   row=241
   col=1
   if len(e)!=0:
      sheet8.write(240, 2, 'Ед.Изм', style=style1)
      for z in e:  
        sheet8.write(row, 2, z, style=style1)
        row+=1 
if len(TrackPointb_6)>5:
   a=TrackPointb_6[0:55]
   b=TrackPointb_6[55:114]
   c=TrackPointb_6[114:173]
   d=TrackPointb_6[173:232]
   e=TrackPointb_6[232:291]
   row=5
   col=0
   for y in a:
     sheet8.write(row, 1, '', style=style1)
     sheet8.write(row, 3, y, style=style)
     row+=1  
   if len(b)!=0:   
     sheet8.write(60, 3, 'Всего', style=style)
     row=61
     col=0 
     for y in b:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1 
   if len(c)!=0:   
     sheet8.write(120, 3, 'Всего', style=style)
     row=121
     col=0 
     for y in c:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1 
  
   if len(d)!=0:   
     sheet8.write(180, 3, 'Всего', style=style)
     row=181
     col=0 
     for y in d:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1 
   if len(e)!=0:   
     sheet8.write(240, 3, 'Всего', style=style)
     row=241
     col=0 
     for y in e:  
       sheet8.write(row, 3, y, style=style)
       sheet8.write(row, 1, '', style=style1)
       row+=1
wb.save('Main_list.xls')
wb=0
#_________________________________________
print ('ver.: 2.09')
print ('Готово, записал все значения в таблицу и сформировал заказ на склад!!!')
print ('Маршруты и накладные сформированны')
print ('Общий заказ в и расчет по маршрутам в Main_list.xls')
print ('Накладные в маршрутах:\nNakladnie-Mar-1.xls\nNakladnie-Mar-2.xls')
print ('Nakladnie-Mar-3.xls\nNakladnie-Mar-4.xls\nNakladnie-Mar-5.xls')
