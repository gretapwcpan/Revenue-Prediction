# -*- coding: utf-8 -*-
"""
Created on Wed Dec 16 14:13:06 2020

@author: gretapan

"""

import pandas as pd
import numpy as np
from datetime import datetime
import math
from io import BytesIO
import matplotlib.pyplot as plt
from styleframe import StyleFrame, Styler, utils
import time
import win32com.client as win32 #to adjust excel automatically

df = pd.read_feather('teradataresult.feather')

working_file = df[['Region','Distributor','End Customer Name','DEBIT',\
                   'Reg Effort','Part Win Date','Design Part Win Status',\
                       'Design Reg Status','Design Item Reg Status','POS Amt',\
                           'Project Units','POS Qty','Design Item','AVG DC',\
                               'AVG DS','AVG SC','Book Part','MAG Desc','Project',\
                                   'DW Amt']]

#since groupby will automatically exclude np.nan
#replace nan with missing as nan is not definable
# =============================================================================
# working_file['DEBIT'].fillna('M', inplace = True)
# working_file['Part Win Date'].fillna('M', inplace = True)
# working_file['POS Amt'].fillna('M', inplace = True)
# working_file['Project Units'].fillna('M', inplace = True)
# =============================================================================

working_file = working_file.fillna('M')

# =============================================================================
# #message type
# 1. Missing Value 
# 2. Below threshold
# =============================================================================

#control points
dr_quote = list(map(lambda i: 2 if (i == 'M') else 0, working_file['DEBIT']))
dr_quote_msg = list(map(lambda i: 'Missing Quote' if (i == 'M') else 'Pass', \
                        working_file['DEBIT']))
dr_debit = list(map(lambda i: 2 if (i == 'N' or i =='M') else 0, working_file['DEBIT']))
dr_debit_msg = list(map(lambda i: 'Missing Debit' if (i == 'N') else ('Pass' if (i == 'Y') else 'Missing Quote'), \
                        working_file['DEBIT']))
#quote_debit = list(map(lambda i: 2 if (i != 'Y') else 0, working_file['DEBIT']))
dr_dw = list(map(lambda i, j: 2 if (i == 'M' or j == 'pending') else 0, \
                 working_file['Part Win Date'], working_file['Design Part Win Status']))
dr_dw_msg = list(map(lambda i, j: 'Missing Design Win' if (i == 'M' or j == 'pending') \
                     else 'Pass', working_file['Part Win Date'], working_file['Design Part Win Status']))

class Controls:
    def __init__(self):
        self.qty_pos_dw = []
        self.amt_pos_dw = []
        self.pm = []
        self.dm = []
        
    def get_qty_pos_dw():
        pointer = 0
        valid_pos_qty = []
        valid_dw_qty = []
        res_pos_dw_qty = []
        while pointer < len(working_file['Distributor']):
            # check valid DW quantity
            if working_file['Part Win Date'][pointer] != 'M' and working_file['Design Part Win Status'][pointer] != 'Pending':
                valid_dw_qty.append(working_file['Project Units'][pointer]/1000)
            else:
                valid_dw_qty.append(np.nan)
            #check valid POS 
            if working_file['POS Qty'][pointer] != 'M':
                valid_pos_qty.append(working_file['POS Qty'][pointer]/1000)
            else:
                valid_pos_qty.append(np.nan)
            pointer += 1
        res_pos_dw_qty = list(map(lambda i, j: i/j, valid_pos_qty, valid_dw_qty))
        scr_pos_dw_qty = list(map(lambda i: 0 if (i > 0.4) else 1, res_pos_dw_qty)) 
        pos_dw_qty_msg = list(map(lambda x,y,z: 'Missing POS or DW' if (math.isnan(x) or math.isnan(y)) \
                                  else ('Below POS DW Quantity ratio 0.4' if z == 1 else 'Pass'),valid_pos_qty, valid_dw_qty,scr_pos_dw_qty))          
        return res_pos_dw_qty, scr_pos_dw_qty, pos_dw_qty_msg
    
    def get_amt_pos_dw():
        pointer = 0
        valid_pos_amt = []
        valid_dw_amt = []
        res_pos_dw_amt = []
        while pointer < len(working_file['Distributor']):
            # check valid DW amt
            if working_file['Part Win Date'][pointer] != 'M' and working_file['Design Part Win Status'][pointer] != 'Pending':
                valid_dw_amt.append(working_file['DW Amt'][pointer]/1000)
            else:
                valid_dw_amt.append(np.nan)
            #check valid POS amt
            if working_file['POS Amt'][pointer] != 'M':
                valid_pos_amt.append(working_file['POS Amt'][pointer]/1000)
            else:
                valid_pos_amt.append(np.nan)
            pointer += 1
        res_pos_dw_amt = list(map(lambda i, j: i/j, valid_pos_amt, valid_dw_amt))
        scr_pos_dw_amt = list(map(lambda i: 0 if (i > 0.4) else 1, res_pos_dw_amt)) 
        pos_dw_amt_msg = list(map(lambda x,y,z: 'Missing POS or DW' if (math.isnan(x) or math.isnan(y)) \
                                  else ('Below POS DW Amount ratio 0.4' if z == 1 else 'Pass'),valid_pos_amt, valid_dw_amt, scr_pos_dw_amt))                    
        return res_pos_dw_amt, scr_pos_dw_amt, pos_dw_amt_msg
    
    def get_pm():
        pointer = 0
        res_PM = []
        res_PM_msg = []
        working_file['AVG SC'].fillna('M', inplace = True)
        while pointer < len(working_file['Distributor']):
            if  working_file['AVG SC'][pointer] != 'M':
                res_PM.append((working_file['AVG DC'][pointer] - \
                               working_file['AVG SC'][pointer])/working_file['AVG DC'][pointer])
            else:
                res_PM.append(np.nan)
            pointer += 1
        scr_PM = list(map(lambda i: 2 if (i < 0.5 or math.isnan(i)) else 0, res_PM))
        res_PM_msg = list(map(lambda x,y: 'Missing PM' if (math.isnan(x)) else \
                              ('PM Below 0.5' if y==2 else 'Pass'), res_PM, scr_PM))
        return res_PM, scr_PM, res_PM_msg
  
    def get_dm():
        pointer = 0
        res_DM = []
        res_DM_msg = []
        scr_DM = []
        working_file['AVG DS'].fillna('M', inplace = True)
        working_file['AVG DC'].fillna('M', inplace = True)
        while pointer < len(working_file['Distributor']):
            if  working_file['AVG DS'][pointer] != 'M':
                res_DM.append((working_file['AVG DS'][pointer] - working_file['AVG DC'][pointer])/working_file['AVG DS'][pointer])
            else:
                res_DM.append(np.nan)
            pointer += 1
        #update DM
        #create margin guideline matrix
        #asia includes ('GC', 'KOR', 'SAP')
        margin_guideline = {('AMR','EXPERT'): 0.25,('EMEA','EXPERT'):0.25, \
                            ('JPN','EXPERT'):0.16, 
                            ('ASIA','EXPERT'):0.15, ('AMR','DEMAND CREATION'): 0.16,\
                                ('EMEA','DEMAND CREATION'): 0.16,
                            ('JPN','DEMAND CREATION'): 0.16, \
                                ('ASIA','DEMAND CREATION'): 0.1,\
                                    ('AMR','FULFILLMENT'): 0.07,
                            ('EMEA','FULFILLMENT'): 0.07, ('JPN','FULFILLMENT'): 0.04, \
                                ('ASIA','FULFILLMENT'): 0.03}
        
        
        for index, element in enumerate(res_DM):
            region = working_file['Region'][index].strip()
            effort = working_file['Reg Effort'][index]
            if region == 'GC' or region =='KOR' or region == 'SAP':
                region = 'ASIA'
            try:
                if element > margin_guideline[(region, effort)]:
                    scr_DM.append(1)
                    res_DM_msg.append('DM higher than guideline')
                elif math.isnan(element):
                    scr_DM.append(1)
                    res_DM_msg.append('Missing DM')
                else:
                    scr_DM.append(0)
                    res_DM_msg.append('Pass')
            except:
                scr_DM.append(1)
                res_DM_msg.append('Missing DM')
                #res_DM_msg.append('unexpected situation')
        return res_DM, scr_DM, res_DM_msg
    
  
pos_dw_amt, scr_pos_dw_qty, pos_dw_qty_msg = Controls.get_qty_pos_dw()
pos_dw_qty, scr_pos_dw_amt, pos_dw_amt_msg = Controls.get_amt_pos_dw()
res_PM, scr_PM, res_PM_msg = Controls.get_pm()
res_DM, scr_DM, res_DM_msg = Controls.get_dm()

# =============================================================================
# #move back to functions for readability
# scr_pos_dw_amt = list(map(lambda i: 0 if (i > 0.4) else 1, amt_pos_dw))
# scr_pos_dw_qty = list(map(lambda i: 0 if (i > 0.4) else 1, qty_pos_dw))     
# scr_res_PM = list(map(lambda i: 2 if (i < 0.5 or math.isnan(i)) else 0, res_PM))
# =============================================================================
#to do: investigate why nan is considered as > 0.5 only in PM's case

SUM = list(map(lambda a,b,c,d,e,f,g: sum([a,b,c,d,e,f,g]), dr_quote, dr_debit, \
               dr_dw, scr_pos_dw_amt, scr_pos_dw_qty, scr_PM, scr_DM))

column_name = ['pos_dw_amt', 'pos_dw_qty', 'PM','DM','dr_quote', 'dr_debit', 'dr_dw', \
               'pos_dw_amt_score','pos_dw_qty_score','PM_score',\
                   'DM_score','SUM','DR_quote message','DR_debit message',\
                       'DR_dwin message','POS_DW_Q_Message','POS_DW_A_Message','DM message', 'PM Message']
column_variable = [pos_dw_amt, pos_dw_qty,  res_PM, res_DM, dr_quote, dr_debit, dr_dw, \
                   scr_pos_dw_amt, scr_pos_dw_qty, scr_PM, \
                       scr_DM, SUM, dr_quote_msg, dr_debit_msg, dr_dw_msg, \
                           pos_dw_qty_msg , pos_dw_amt_msg, res_DM_msg, res_PM_msg]

#insert testdata into blank dataframe
for index, element in enumerate(column_name):
    working_file[element] = column_variable[index]

# =============================================================================
# Error_Message = []
# def error_message():
#     linetext = ''
#     for index, element in enumerate(working_file['region']):
#         if working_file['DEBIT'][index] == 'M':
#             linetext += 'Missing Quote'
#         elif 
#         
# =============================================================================

#Error message count
error_msg = pd.DataFrame()
column_name = ['DR_quote message','DR_debit message',\
                       'DR_dwin message','POS_DW_Q_Message','POS_DW_A_Message','DM message', 'PM Message']
column_var = [dr_quote_msg, dr_debit_msg, dr_dw_msg, pos_dw_qty_msg , pos_dw_amt_msg, res_DM_msg, res_PM_msg]

for index, element in enumerate(column_name):
    error_msg[element] = column_var[index]


#temporary check file
#working_file.to_csv('check_{}.csv'.format(datetime.today().strftime('%Y_%m_%d')))


# =============================================================================
# # An "interface" to matplotlib.axes.Axes.hist() method
# n, bins, patches = plt.hist(x=SUM, bins='auto', color='#0504aa',
#                             alpha=0.7, rwidth=0.85)
# plt.grid(axis='y', alpha=0.75)
# plt.xlabel('Scores')
# plt.ylabel('Counts')
# plt.title('Total Score Histogram')
# plt.text(23, 45, r'$\mu=15, b=3$')
# maxfreq = n.max()
# # Set a clean upper y-axis limit.
# plt.ylim(ymax=np.ceil(maxfreq / 10) * 10 if maxfreq % 10 else maxfreq + 10)
# =============================================================================

#do not replace vanila writer object with styleframe excel writer object as engine can't be revised in styleframe
outputname = 'output{}.xlsx'.format(datetime.today().strftime('%Y_%m_%d'))
writer = pd.ExcelWriter(outputname, engine='xlsxwriter')
#ew = StyleFrame.ExcelWriter(outputname)
#stat = StyleFrame(pd.DataFrame())
stat = pd.DataFrame()
stat.to_excel(writer, sheet_name = 'Descriptive Statistics')
worksheet = writer.sheets['Descriptive Statistics']

img = working_file['SUM'].hist(bins = 15) #add granularity
imagedata = BytesIO()
fig = plt.figure()
img.figure.savefig(imagedata)

image_width = 128.0
image_height = 40.0

cell_width = 64.0
cell_height = 20.0
x_scale = cell_width/image_width
y_scale = cell_height/image_height
worksheet.write('B2','Distribution of final score')
#worksheet.cell(row=2, column=2, value='Distribution of final score')
worksheet.insert_image('B3','',{'image_data':imagedata, 'x_scale': x_scale, 'y_scale': y_scale})

#styleframe has been suppressed given that its effect isn't good enough
def Calerror_page():
    R = 15
    C = 1
    worksheet.write('B14','Statistics of Error Message')
    #ew = StyleFrame.ExcelWriter(writer)
    for msg in column_name:
        error_msg_page = working_file[msg].groupby(working_file[msg]).describe()[['count']]
        error_msg_page['percent'] = (error_msg_page['count'] / len(working_file[msg])).apply(lambda x : str(round(x*100,2)) + '%')
        error_msg_page.append({msg: 'Total', 'count': len(working_file[msg]), 'percent': 100}, ignore_index=True)
        #sf = StyleFrame(error_msg_page)
        #sf.to_excel(ew,sheet_name='Descriptive Statistics',startrow = R , startcol= C) 
        error_msg_page.to_excel(writer,sheet_name='Descriptive Statistics',startrow = R , startcol= C)
        R += error_msg_page.shape[0] + 2
    #ew.save()

Calerror_page()

#Missing Value Smmary-By distributor
def CalMis(view, startcolumn, charttype):
    msg = ['DR_quote message','DR_debit message','DR_dwin message','PM Message', 'DM message','POS_DW_A_Message','POS_DW_Q_Message']
    threshold = ['Missing Quote','Missing Debit','Missing Design Win','Missing PM','Missing DM','Missing POS or DW','Missing POS or DW']
    R = 2
    if charttype == 'pie':
        for index, element in enumerate(msg):
            df = working_file[element][working_file[element] == threshold[index]].groupby(view).describe()
            plot = df.plot.pie(y='count', figsize=(5, 5), autopct='%1.1f%%', legend = False)
            imagedata = BytesIO()
            #fig = plt.figure()
            plot.figure.savefig(imagedata)
            
            image_width = 112.0
            image_height = 40.0
            
            cell_width = 64.0
            cell_height = 20.0
            x_scale = cell_width/image_width
            y_scale = cell_height/image_height
            worksheet.write('{}{}'.format(startcolumn, R), threshold[index])
            #worksheet.cell(row=2, column=2, value='Distribution of final score')
            worksheet.insert_image('{}{}'.format(startcolumn, R+1),'',{'image_data':imagedata, 'x_scale': x_scale, 'y_scale': y_scale})
            R += 14
    else:
        for index, element in enumerate(msg):
            df = working_file[element][working_file[element] == threshold[index]].groupby(view).describe()
            df = df.sort_values(by='count', ascending=True)
            plot = df.plot(kind = 'barh', y='count', figsize=(6, 4), fontsize=7)
            imagedata = BytesIO()
            #fig = plt.figure()
            #bbox_inches='tight'-->expand the region to include xticks
            plot.figure.savefig(imagedata, bbox_inches='tight')
            
            image_width = 112.0
            image_height = 40.0
            
            cell_width = 64.0
            cell_height = 20.0
            x_scale = cell_width/image_width
            y_scale = cell_height/image_height
            worksheet.write('{}{}'.format(startcolumn, R), threshold[index])
            #worksheet.cell(row=2, column=2, value='Distribution of final score')
            worksheet.insert_image('{}{}'.format(startcolumn, R+1),'',{'image_data':imagedata \
                                                                       ,'x_scale': x_scale, 'y_scale': y_scale\
                                                                           })
            R += 14

#Below Threshold Summary-By region or By distributor
def CalTH(view, startcolumn, charttype):
    msg = ['PM Message', 'DM message','POS_DW_A_Message','POS_DW_Q_Message']
    threshold = ['PM Below 0.5','DM higher than guideline','Below POS DW Amount ratio 0.4','Below POS DW Quantity ratio 0.4']
    R = 2
    if charttype == 'pie':
        for index, element in enumerate(msg):
            df = working_file[element][working_file[element] == threshold[index]].groupby(view).describe()
            plot = df.plot.pie(y='count', figsize=(5, 5), autopct='%1.1f%%', legend = False)
            imagedata = BytesIO()
            #fig = plt.figure()
            plot.figure.savefig(imagedata)
            
            image_width = 112.0
            image_height = 40.0
            
            cell_width = 64.0
            cell_height = 20.0
            x_scale = cell_width/image_width
            y_scale = cell_height/image_height
            worksheet.write('{}{}'.format(startcolumn, R), threshold[index])
            #worksheet.cell(row=2, column=2, value='Distribution of final score')
            worksheet.insert_image('{}{}'.format(startcolumn, R+1),'',{'image_data':imagedata, 'x_scale': x_scale, 'y_scale': y_scale})
            R += 14
    else:
        for index, element in enumerate(msg):
            df = working_file[element][working_file[element] == threshold[index]].groupby(view).describe()
            df = df.sort_values(by='count', ascending=True)
            plot = df.plot(kind = 'barh', y='count', figsize=(6, 4), fontsize=8)
            imagedata = BytesIO()
            #fig = plt.figure()
            #bbox_inches='tight'-->expand the region to include xticks
            plot.figure.savefig(imagedata, bbox_inches='tight')
            
            image_width = 112.0
            image_height = 40.0
            
            cell_width = 64.0
            cell_height = 20.0
            x_scale = cell_width/image_width
            y_scale = cell_height/image_height
            worksheet.write('{}{}'.format(startcolumn, R), threshold[index])
            #worksheet.cell(row=2, column=2, value='Distribution of final score')
            worksheet.insert_image('{}{}'.format(startcolumn, R+1),'',{'image_data':imagedata \
                                                                       ,'x_scale': x_scale, 'y_scale': y_scale\
                                                                           })
            R += 14

    
CalTH(working_file['Region'],'F', 'pie')
CalTH(working_file['Distributor'],'I', 'bar')
CalMis(working_file['Region'], 'M', 'pie')
CalMis(working_file['Distributor'], 'Q', 'bar')

#working_file = StyleFrame(working_file)
# unable to use engine xlsxwriter to solve IllegalCharacterError in style frame


#styling raw data sheet
# Get the xlsxwriter workbook and worksheet objects.
working_file.to_excel(writer, sheet_name = 'Complete Raw Data')
workbook  = writer.book
worksheet = writer.sheets['Complete Raw Data']
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})

for col_num, value in enumerate(working_file.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)
    

writer.save()

time.sleep(0.8)
#pull path to open excel is required
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open('C:\\Users\\nxf33342\\Desktop\\disti_DR_project\\' + outputname)
ws = wb.Worksheets("Descriptive Statistics")
wst = wb.Worksheets("Complete Raw Data")
ws.Columns.AutoFit()
wst.Columns.AutoFit()
wb.Save()
excel.Application.Quit()
