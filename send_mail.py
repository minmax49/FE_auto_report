# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 09:24:06 2020

@author: nguyenquangminh3
;hoang.le.8@fecredit.com.
dang.tran.3@fecredit.com.vn;hoang.le.8@fecredit.com.vn' 
"""

#import os
#os.chdir('C:/Users/nguyenquangminh3/projects/Mr_Dang_mail/') 
from datetime import datetime
import win32com.client as win32  
import pickle

#%% step 1 : mail setup
# =============================================================================
# mail setup
# =============================================================================

def send(mail_name='minh.nguyen.50@fecredit.com.vn',
         main_path = 'C:/Users/nguyenquangminh3/projects/Miss_Uyen_report',
         mail_to ='minh.nguyen.50@fecredit.com.vn;'):
    
    
    mail = win32.Dispatch('outlook.application').CreateItem(0)
    # select acount to send mail:
    mail.SentOnBehalfOfName = mail_name
    
    # mail to :
    mail.To = mail_to
    #mail.To = 'minh.nguyen.50@fecredit.com.vn;'
    # mail Subject = CRC Field Operation Report - Jan'20 MTD
    mail.Subject = "CRC Field Operation Report - {} MTD".format(
        datetime.now().strftime("%h'%y"))
    
    # Attachments files with id 
    path1 = main_path + '\\image\\temp1.png' 
    attachment = mail.Attachments.Add(path1 )
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
    
    path2 = main_path + '\\image\\temp2.png' 
    attachment = mail.Attachments.Add(path2)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId2")
    
    path3 = main_path + '\\image\\temp3.png' 
    attachment = mail.Attachments.Add(path3)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId3")
    
    path4 = main_path + '\\image\\temp4.png' 
    attachment = mail.Attachments.Add(path4)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId4")
    
    path5 = main_path + '\\image\\temp5.png' 
    attachment = mail.Attachments.Add(path5)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId5")
    
    path6 = main_path + '\\image\\temp6.png' 
    attachment = mail.Attachments.Add(path6)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId6")
    

    
    
    #%% step 2 : mail contents
    # =============================================================================
    # mail HTMLBody
    # =============================================================================
    
    with open('comment.pickle', 'rb') as handle:
        comment = pickle.load(handle)
    
    
    html_header = """
    <p style="font-size:18px"><b> Dear ALL,</b> <br>
    Kindly help to refer to CRC Field Operation Report {0} </b><br>
    <h4>Note </h4>
    <ul style="font-size:18px"> 
        <li > %RESOLVED_MTD of FC-CRC overall is GAP <b style="color:'red'">{1}%</b> compare with Dec2019 and GAP  <b style="color:'red'">{2}%</b> compare with Nov2019 </li>
        <li > %RESOLVED_MTD of FCS-CRC overall is GAP <b style="color:'red'">{3}%</b> compare with Dec2019 and GAP  <b style="color:'red'">{4}%</b> compare with Nov2019</li>
    </ul> 
    """.format(datetime.now().strftime("%h'%y"), 
        comment['FC-CRC1'],
        comment['FC-CRC2'],
        comment['FCS-CRC1'],
        comment['FCS-CRC2'])
    
    
    html_main = """
    <div style="width:100%;float:left">
        <p style="margin:15;float:left"><img src="cid:MyId1" > </p>
        <p style="margin:15;float:left"><img src="cid:MyId2" > </p>
    </div>
    
    <div style="width:100%;float:left">
        <h3 style="color:'blue'"> BUCKET </h3> 
        
        <p style="margin:15;float:left"><img src="cid:MyId3" > </p>
    </div>
    
    <div style="width:100%;float:left">
        <h3 style="color:'blue'"> TARGET </h3> 
        <p style="margin:15;float:left"><img src="cid:MyId4" > </p>
        <p style="margin:15;float:left"><img src="cid:MyId5"> </p>
    </div>
    
    <div style="width:100%;float:left">
        <h3 style="color:'blue'"> Resource Analysis </h3> 
        <p style="margin:15;float:left"><img src="cid:MyId6"> </p>
    
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
