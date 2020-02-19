create or replace procedure COL_SP_CREDIT_CARD is
 v_tmp number; 
 v_tmp2 number:=0;
 
begin

  -------------create table current monthly table just for card-------------------
  v_tmp:=0;
  select /*+first_rows(1) parallel(4)*/count(1)
  into v_tmp
  from cs_case_info_his_crc_cur_mon t
  where t.bank_date = trunc(sysdate);
  
  if v_tmp > 0 then
  
  delete from cs_case_info_his_crc_cur_mon t where t.bank_date = trunc(sysdate);
  commit;
  end if;
  
  v_tmp :=0;
  while v_tmp < 1 loop
      select count(1)
      into v_tmp
      from ODS.SCHEDULE_IN_JOB_LOG where trunc(finished_date) = trunc(sysdate)
      and job_name='CASE_HISTORY_INSERT_DAILY'
      and job_step_done='CS_CASE_INFO_TODAY';
   dbms_lock.sleep(120);
  end loop;
  --Add 2019SEP04 
  v_tmp2 :=0;
  while v_tmp2 < 1 loop
      select count(1)
      into v_tmp2
      from sdm.sdm_log_etl_daily_monitor a 
      where a.TABLE_NAME ='SDM.SDM_COL_BALANCE_CRC'
      and trunc(a.START_TIME) = trunc(sysdate) and a.STATUS = 'COMPLETED' ;
   dbms_lock.sleep(120);
  end loop;
  --End 
  -- 2 dieu kien thoa --
  if v_tmp > 0 and v_tmp2 >0 then
    

  
  insert /*+ enable_parallel_dml parallel(4) */ into cs_case_info_his_crc_cur_mon  
   select /*+ FULL(T) parallel(4)*/ t.appl_id, t.queue, t.allocation_date, trunc(sysdate) bank_date, t.unit_id, t.dpd, t.cycle_due, t.min_amount_due,
t.total_amt_due, t.total_curr_due, t.payment_due_date, t.last_statement_date, t.next_statement_date, t.amount_outstanding,
t.curr_charges_outstanding, t.curr_pos_outstanding, t.bucket, t.amount_overdue, t.interest_overdue, t.interest_outstanding,
t.curr_int_outstanding, t.card_type, t.card_catg,t.memo_balance, t.credit_limit, t.other_charges, t.bom_bucket,t.casetype,
t.branch_code, t.cust_name,
t.hi_billed, t.bucket1, t.bucket2,
t.bucket3, t.bucket4, t.bucket5, t.bucket6,t.bucket7, t.bucket8, t.bucket9, t.bucket10, t.card_expiry_date,
t.product_type, t.charge_off_flag, t.charge_off_date, t.over_limit, t.charge_off_reason_code, t.memo_permanent,
t.last_cash_advance_amt, t.last_cash_advance_date, t.last_credit_limit, t.last_credit_limit_date, t.last_delinquency_date,
t.last_payment_amt, t.last_payment_date,t.last_purch_amt, t.last_purch_date, t.last_statement_balance, t.late_fees,
t.ltd_payments, t.ltd_purchases, t.over_limit_flag, t.credit_limit_date, t.hi_billed_date, t.queue_date,
t.account_block_code1, t.account_block_code1_date, t.account_block_code2, t.account_block_code2_date,
t.ltd_retail_purchases_number, t.ltd_retail_purchases_amount, t.ltd_cash_purchases_number, t.ltd_cash_purchases_amount,
t.opened_date, t.currency, t.cash_withdrawal_limit, t.cash_balance, t.cash_available, t.name_on_card, t.cust_id,
t.card_number, t.product, t.grace_days, t.delinquency_string, t.principle_outstanding, t.queue_type
  FROM Dwcollmain.Cs_Case_Info_today t
  where t.branch_code = '1';
  
  commit;

  end if;

execute immediate 'begin
                   dbadw.call_common_credit_card_jobs;
                   end;';
                   
 /* execute immediate 'begin
                     RUN_COL_SP_CE_DAILY_DTL_CC(trunc(Sysdate));
                     end;';
           
  execute immediate 'begin
                     common.CRC_MAI_REPORT_DELI;
                     end;';    
                     
   execute immediate 'begin
                     common.CRC_MAI_REPORT_PREDUE;
                     end;';     
                                                        
  execute immediate 'begin
                     common.COL_SP_CRC_AUTO_PRE_NEW;
                     end;';

  execute immediate 'begin
                     common.COL_SP_CRC_AUTO_BOM;
                     end;';
                                         
  execute immediate 'begin
                     Common.SP_CRC_GLX_PRE_NEW;
                     end;'; */
              
/*  If trunc(sysdate) = '09Sep2017' then
      execute immediate 'begin
                     common.COL_SP_CRC_AUTO_PRE_NEW_M;
                    end;';
      execute immediate 'begin
                     common.COL_SP_CRC_AUTO_BOM_M;
                    end;';
    End If;*/

  /*execute immediate 'begin
                     common.col_sp_crc_auto_pre;
                    end;';*/


end COL_SP_CREDIT_CARD;
