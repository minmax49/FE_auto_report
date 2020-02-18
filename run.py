# -*- coding: utf-8 -*-
"""
Created on Tue Jan 14 13:52:18 2020

@author: nguyenquangminh3
"""

import cx_Oracle
import datetime
import pandas as pd
import time 
import refresh_excel

con = cx_Oracle.connect(user='common',password='Nwpass_Rkcol_0819',
                                dsn='dwproddc',encoding = 'utf-8',
                                nencoding = 'utf-8')

def main():
    print(datetime.datetime.now().strftime('%H:%M'))
    #todayb = datetime.datetime.today().strftime('%d%b%Y')
    #print('run_store')
    #cur = con.cursor()   
    #cur.execute("begin RUN_COL_SP_CE_DAILY_DTL_CC(trunc(sysdate));end;")
    try:
        while True:    
            today = datetime.datetime.today().strftime('%Y/%m/%d')
            sql =  """select count(*) count
            from csa_log_report_check t
            where t.run_date >= TO_TIMESTAMP('{0}', 'yyyy/mm/dd') 
            and t.report_name = 'CRC Collections Performance'""".format(today)
        
            df = pd.read_sql(sql,con)
            count = df['COUNT'].iloc[0]
            
            if count > 0:
                h_start = datetime.datetime.now().strftime('%d/%m/%y %H:%M')
                start = time.time()/60
                print("=="*20)
                print('start running')
        # =========================================================================
        #       set path
        # =========================================================================
                refresh_excel.main()
                print("-------------------------------------")
                end = round(time.time()/60 - start,2)
                print(end , 'completed all')
                
                h_end = datetime.datetime.now().strftime('%d/%m/%y %H:%M')  
                with open ('log.txt' , 'w') as file:
                    file.write('from {} to {} completd in {} ninute'.format(h_start,h_end,end) )
                break
            time.sleep(100)
    except Exception as e:
        print(e)
        time.sleep(10)
        
    print(datetime.datetime.now().strftime('%H:%M'))
    time.sleep(10)

if __name__ =="__main__":
    main()
    
