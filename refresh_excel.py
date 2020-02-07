# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 10:06:00 2020

@author: nguyenquangminh3


Sub aa()
Sheets("EA-Provinces").Range("AR11").Value = 3
        
   Call MAP_DATA
   Call Fill_color
End Sub

# --- 
Còn thiếu refresh các sheet dashboard

"""

import time 
import xlwings as xw
import excel2img
import pandas as pd
import os
import datetime
import shutil
import pickle

import send_mail



def refresh_pivot(wb , sheet='Overview', pivot_name='PivotTable2'):
    """
    Refresh the selected pivot data

    Parameters
    ----------
    sheet : str , name of sheet
         The default is 'Overview'.
    pivot_name : str, name of pivot
         The default is 'PivotTable2'.
    """
    wb.sheets[sheet].select()
    wb.api.ActiveSheet.PivotTables(pivot_name).PivotCache().refresh()
    
    

def update_rundate(wb , sheet='Overview', pivot_name='PivotTable2'):
    today = datetime.datetime.today().strftime('%d') 
    #wb.sheets['Overview'].range('B2').value = '09'
    
    wb.sheets[sheet].select()
    x = wb.api.ActiveSheet.PivotTables(pivot_name).PivotFields("RUN_DATE")
    x.ClearAllFilters()
    x.EnableMultiplePageItems = False
    x.CurrentPage = today



def export_img_func(path, name_img,sheet_name,range_cell ):
    excel2img.export_img(path, name_img,None,
                         "'{}'!{}".format(sheet_name,range_cell))
    



    

def main(main_path = 'C:/Users/nguyenquangminh3/projects/Miss_Uyen_report',
        template_path= 'template.xlsm', 
        save_path = None,
        mail_name = 'minh.nguyen.50@fecredit.com.vn',
        mail_to = 'minh.nguyen.50@fecredit.com.vn'):
    
    """
    main process ,  whit 3 steps : 
        step 1 : define paths , setup things to do job
        step 2 : refresh and shoot , images save to folder : ~/image 
        step 3 : send_mail  , from send_mail.py

    Parameters
    ----------
    main_path : TYPE, optional
        DESCRIPTION. The default is 'C:/Users/nguyenquangminh3/projects/Mr_Dang_mail'.
    template_path : TYPE, optional
        DESCRIPTION. The default is 'template.xlsb'.
    save_path : TYPE, optional
        DESCRIPTION. The default is None.
    mail_name : TYPE, optional
        DESCRIPTION. The default is 'minh.nguyen.50@fecredit.com.vn'.

    """

# =============================================================================
#   step 1 : define paths 
# =============================================================================
    start = time.time()/60
    # remove last '/'
    if main_path[-1] == '/':
        main_path = main_path[:-1]
    
    #set root path
    os.chdir(main_path) 
    temp_excel_path = main_path+ '/' + template_path
    
    # copy template 
    today = datetime.datetime.today().strftime('%Y%m_%d') 
    excel_path = main_path + '/excel/Field Operation Report {}.xlsm'.format(today)
   
    if save_path != None:
        if save_path[-1] == '/':
            save_path = save_path[:-1]
        excel_save_path = save_path + '/Field Operation Report {}.xlsm'.format(today)
        
        
    shutil.copy(temp_excel_path,excel_path)
    refresh_df = pd.read_excel('setup.xlsx',sheet_name='refresh') 
    refresh_df.dropna(inplace=True)
    
    excel_img_df = pd.read_excel('setup.xlsx',sheet_name='excel_img') 
    excel_img_df.dropna(inplace=True)
    
    rundate_df = pd.read_excel('setup.xlsx',sheet_name='rundate') 
    rundate_df.dropna(inplace=True)
# =============================================================================
#   step 2 : refresh and shoot 
# =============================================================================
    # Set hidden excel
    #xw.App(visible=False)
    
    # open workbook.
    wb = xw.Book(excel_path)
    """
    connects = wb.api.Connections
    for i in range(1,connects.Count):
        connects.Item(i).OLEDBConnection.Connection = "OLEDB;Provider=OraOLEDB.Oracle.1;Password=Nwpass_Rkcol_0819;Persist Security Info=True;User ID=Common;Data Source=dwproddc;"
    wb.api.RefreshAll()
    """
    # refresh pivots
    for i in range(len(refresh_df)):
        line = refresh_df.iloc[i]
        refresh_pivot(wb = wb,sheet=line['sheet_name'], pivot_name=line['pivot_name'])
        
    # set rundate to day 
    for i in range(len(rundate_df)):
        line = refresh_df.iloc[i]
        update_rundate(wb = wb,sheet=line['sheet_name'], pivot_name=line['pivot_name'])
    
    #call macro
    
    map_data = wb.macro('Module2.MAP_DATA')
    map_data()
    fill_color = wb.macro('Module1.Fill_color')
    fill_color()
    
    # save comment
    LR_b3 = {'FC-CRC1': round(wb.sheets['Note'].range("B1").value*100,1),
         'FC-CRC2': round(wb.sheets['Note'].range("B2").value*100,1),
         'FCS-CRC1': round(wb.sheets['Note'].range("B3").value*100,1),
         'FCS-CRC2': round(wb.sheets['Note'].range("B4").value*100,1),
         'excel_path':excel_save_path
        }
    with open('comment.pickle', 'wb') as handle:
        pickle.dump(LR_b3, handle, protocol=pickle.HIGHEST_PROTOCOL)
        
    
    for sheet in excel_img_df.sheet_name.unique():
        sht = wb.sheets[sheet]
        sht.autofit('c')
    # save file and quit
    
    
    wb.save()
    wb.app.quit()    
     
    if save_path != None:
        shutil.copy(excel_path,excel_save_path)
    print(time.time()/60 - start, 'completed refresh')
    
    # shoot images

    if not os.path.exists(main_path+ '/image'):
        os.mkdir(main_path+ '/image')
    else:
        shutil.rmtree(main_path+ '/image')
        os.mkdir(main_path+ '/image')

    
    print("shooting now")
    for i in range(len(excel_img_df)):
        temp = excel_img_df.iloc[i]
        export_img_func(excel_path, temp['name_img'],temp['sheet_name'],temp['range_cell'] )
   
    print(time.time()/60 - start , 'completed shooting images')
    
   
# =============================================================================
#   step 3 : send_mail
# =============================================================================
    mail_list = pd.read_excel('setup.xlsx',sheet_name='main_setup')['mail_to']
    mail_list.dropna(inplace=True)
    mail_list = str(mail_list.tolist()).replace("'","").replace("[","").replace("]","").replace(",",";")

    
    send_mail.send(mail_name=mail_name,
                   main_path= main_path,
                   mail_to= mail_list)
   
    print(time.time()/60 - start, 'completed sending mail')

if __name__ == "__main__":
    try:
        main(save_path = 'T:/DeptData/_CollA_Dev_Report')
    except Exception as e:
        print(e)
        time.sleep(100)
    
  
