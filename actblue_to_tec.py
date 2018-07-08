'''
Created on Jan 2, 2018

@author: Matt
'''
import csv
from datetime import datetime
import tkFileDialog
import os

actblue_filename = 'actblue 2017-07-01 to 2017-12-31.csv'

actblue_filename = tkFileDialog.askopenfilename(title='ActBlue Export File',filetype=(("csv files","*.csv"),))

if actblue_filename:
    tec_contrib_filename = tkFileDialog.asksaveasfilename(title='TEC Contribution File',filetype=(("csv files","*.csv"),))
    tec_expense_filename = tkFileDialog.asksaveasfilename(title='TEC Expense File',filetype=(("csv files","*.csv"),))
    
    if os.path.splitext(tec_contrib_filename)[1].lower() != '.csv':
        tec_contrib_filename += '.csv'
    if os.path.splitext(tec_expense_filename)[1].lower() != '.csv':
        tec_expense_filename += '.csv'
    
    
    
    
    if tec_contrib_filename and tec_expense_filename:
    
        actblue_items = []
        
        with open(actblue_filename,'rb') as actblue_fin:
            actblue_reader = csv.reader(actblue_fin)
            headings = actblue_reader.next()
            for row in actblue_reader:
                actblue_items += [dict(zip(headings,row))]
        
        tec_contrib_headings = ['#Rec_Type','Form_Type','Item_ID','Entity_Cd','Ctrib_NamL','Ctrib_NamF','Ctrib_NamT','Ctrib_NamS',
                                'Ctrib_Adr1','Ctrib_Adr2','Ctrib_City','Ctrib_StCd','Ctrib_ZIP4','Ctrib_CtryCD',
                                'OS_PAC_CB','OS_PAC_FEC','Ctrib_Date','Ctrib_Amt','Ctrib_Dscr','Employer','Occup','Job_Title','Spous_Law','Parent1','Parent2']
        
        tec_expense_headings = ['#Rec_Type','Form_Type','Item_ID','Entity_Cd','Payee_NamL','Payee_NamF','Payee_NamT','Payee_NamS',
                                'Payee_Adr1','Payee_Adr2','Payee_City','Payee_StCd','Payee_ZIP4','Payee_CtryCD',
                                'Expn_Date','Expn_Amt','Expn_Dscr','ExpCntr_YN', 'Reimbur_CB', 
                                'Cand_NamL', 'Cand_NamF', 'Cand_NamT', 'Cand_NamS', 
                                'OffHldCd', 'OffHldNam', 'OffHldDist','OffHldPlace', 'OffSeekCd', 'OffseekNam', 'OffseekDist','OffSeekPlace', 
                                'BakRef_ID', 'ExpnCorp_YN', 'Trvl_CB', 'Trvl_NamL', 'Trvl_NamF', 'Trvl_NamT', 'Trvl_NamS', 'Tran_Type','Tran_Descr', 
                                'Dpt_City', 'Dpt_Date', 'Arv_City', 'Arv_Date', 'Trvl_Purp', 'Trvl_BakRef', 'Expn_CatgOth', 'Expn_Catg', 'Aus_Living_Exp_CB']
        
        with open(tec_contrib_filename, 'wb') as tec_contrib_fout:
            with open(tec_expense_filename, 'wb') as tec_expense_fout:
                tec_contrib_writer = csv.writer(tec_contrib_fout)
                tec_expense_writer = csv.writer(tec_expense_fout)
                 
                tec_contrib_writer.writerow(tec_contrib_headings)
                tec_expense_writer.writerow(tec_expense_headings)
                
                for item in actblue_items:
                    zip_code = item['Donor ZIP']
                    if len(zip_code) < 5:
                        zip_code = '0'*(5-len(zip_code))+zip_code
                    item_date = datetime.strptime(item['Date'],'%Y-%m-%d %H:%M:%S')
                    item_date_str = item_date.strftime('%Y%m%d')
                    tec_contrib_row = ['RCPT','A1','','I',item['Donor Last Name'],
                                       item['Donor First Name'],'','',item['Donor Addr1'],
                                       item['Donor Addr2'],item['Donor City'],item['Donor State'],
                                       zip_code,'USA','','',item_date_str,item['Amount'],
                                       'ActBlue Donation',item['Donor Employer'],
                                       item['Donor Occupation']]
                    tec_contrib_writer.writerow(tec_contrib_row)
                    
                    tec_expense_row = ['EXPN','F1','','E','ActBlue Texas','','','',
                                       'PO Box 382110','','Cambridge','MA','02238-2110',
                                       'USA',item_date_str,item['Fee'],'ActBlue Service Fee',
                                       'Y']+['']*29+['FEES','']
                    tec_expense_writer.writerow(tec_expense_row)
            