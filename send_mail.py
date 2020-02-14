# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 09:24:06 2020
@author: nguyenquangminh3
"""

#import os
#os.chdir('C:/Users/nguyenquangminh3/projects/Mr_Dang_mail/') 
import datetime
import win32com.client as win32  
import pickle
import pandas as pd


#%% step 1 : mail setup
# =============================================================================
# mail setup
# =============================================================================

def send(mail_name='minh.nguyen.50@fecredit.com.vn',
         main_path = 'C:/Users/nguyenquangminh3/projects/CRC_report_01',
         mail_to ='minh.nguyen.50@fecredit.com.vn;',
         mail_cc = ""):
    
    
    mail = win32.Dispatch('outlook.application').CreateItem(0)
    # select acount to send mail:
    mail.SentOnBehalfOfName = mail_name
    
    # mail to :
    mail.To = mail_to
    if mail_cc != "":
        mail.CC = mail_cc
    #mail.To = 'minh.nguyen.50@fecredit.com.vn;'
    # mail Subject = CRC Field Operation Report - Jan'20 MTD
    mail.Subject = "[AUTO-MAIL] CRC Collections Performance Snapshot - {} ___TEST".format(
        datetime.datetime.now().strftime("%h'%y"))
    
    # Attachments files with id 
    excel_img_df = pd.read_excel(main_path +'/setup.xlsx',sheet_name='excel_img') 
    excel_img_df.dropna(subset = ['sheet_name'],inplace=True)
    
    for i in range(len(excel_img_df)):
        path1 = main_path + '/'+excel_img_df.name_img[i]
        #print(main_path + excel_img_df.name_img[i])
        attachment = mail.Attachments.Add(path1 )
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId{}".format(i+1))
        
    #%% step 2 : mail contents
    # =============================================================================
    # mail HTMLBody
    # =============================================================================
    #lastmonth = (datetime.datetime.now().replace(day =  1 ) - datetime.timedelta(days=7)).strftime("%b%Y")
    #last2month = (datetime.datetime.now().replace(day =  1 )  - datetime.timedelta(days=32)).strftime("%b%Y")

    with open('comment.pickle', 'rb') as handle:
        comment = pickle.load(handle)
        
        
    html_header = """
    <p style="font-size:18px;color:'#246f83'"><b> Dear all,</b> <br>
    I would like to update you on the Collections Performance daily in {0} of CRC with item below: </b><br>
    
    <ul style="font-size:18px;color:'#246f83'"> 
        <li>I.Roll Forward </li>
        <li>II. Net Flow</li>
        <li>III. Resolved & Target Collected Amount</li>
        <li>IV. Entry rate 1DPD & 10 DPD</li>  
        </ul> 
    """.format(datetime.datetime.now().strftime("%b'%Y"))
    
    html_main = """
    <div style="width:100%;float:left">
         <h3 style="color:'#21abcd'"> I/ ROLL FORWARD REPORT DAILY: (RF_AMT) </h3>
        <p style="margin:15;float:left"><img src="cid:MyId1" > </p>
        <p style="margin:15;float:left"><img src="cid:MyId2" > </p>
    </div>
    
    <div style="width:100%;float:left">
        <h3 style="color:'#21abcd'"> II/ NET FLOW: </h3> 
        
        <p style="margin:15;float:left"><img src="cid:MyId3" > </p>
    </div>
    
     <div style="width:100%;float:left">
        <h3 style="color:'#21abcd'"> III. RESOLVED & TARGET COLLECTED AMOUNT</h3> 
        <h4 style="color:'#21abcd'"> PRE-DELINQUENT: </h4>
        <p style="margin:15;float:left"><img src="cid:MyId4"> </p>
        <h4 style="color:'#21abcd'"> DELINQUENT: </h4>
        <p style="margin:15;float:left"><img src="cid:MyId5"> </p>
    
    </div>
    <div style="width:100%;float:left">
        <h3 style="color:'#21abcd'"> IV. ENTRY RATE </h3> 
        <h4 style="color:'#21abcd'"> Entry 1 DPD</h4>
         <p style="margin:15;float:left"><img src="cid:MyId6"> </p>
        <h4 style="color:'#21abcd'"> Entry 10 DPD</h4>
         <p style="margin:15;float:left"><img src="cid:MyId7"> </p>
    </div>

    <div>
        <b>Link: {0} </b><br>
        <a href ='{0}'>Go to excel file  </a> <br>
        
        <p><b> Best regards,</b> <br>
        <b>COLLECTION MIS. </b></p>
    </div>
    """.format(comment['excel_path'])
    
    mail.HTMLBody = html_header + html_main
    #mail.Send()
    mail.Display()
if __name__ == "__main__":
    send()
