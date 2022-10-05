from cgi import test
from statistics import stdev
from function_file import *
from openpyxl import Workbook
import win32com.client as win32
import pythoncom
import sys
import os
from sklearn.tree import DecisionTreeRegressor 
import scipy.stats as st
from scipy.stats import norm
from math import sqrt
import pandas as pd
import numpy as np
import re
import datetime
from Ml_model_sscalc import *

# define our Application Events
#The event handlers container

class wsEvents_filterPN:
    def OnClick(self,*args):
        CurrentValueVar = readCurrentValue(ws_1)
        # code to see the change coz of part numberfield:
        InputChangeValidate(ws_3,ws_1,data1,np,'PartNumber',CurrentValueVar,0)

        
class wsEvents_filterPlantDepo:
    def OnClick(self,*args):
        CurrentValueVar = readCurrentValue(ws_1)
        #code to see the change initiated by plant/depo:
        InputChangeValidate(ws_3,ws_1,data1,np,'DepoPlant',CurrentValueVar,0)


class wsEvents_filterregion:
    def OnClick(self,*args):
        CurrentValueVar = readCurrentValue(ws_1)
        #code to see the change initiated by Region:
        InputChangeValidate(ws_3,ws_1,data1,np,'Region',CurrentValueVar,0)

class wsEvents_filtershipregion:
    def OnClick(self,*args):
        CurrentValueVar = readCurrentValue(ws_1)
        #code to see the change initiated by Ship Region:
        InputChangeValidate(ws_3,ws_1,data1,np,'Ship_Region',CurrentValueVar,0)
class wsEvents_filtercus:
    def OnClick(self,*args):
        CurrentValueVar = readCurrentValue(ws_1)
        #code to see the change initiated by customer filter
        InputChangeValidate(ws_3,ws_1,data1,np,'Customer',CurrentValueVar,0)
class wsEvents_deleterecords:
    def OnClick(self,*args):
        ws_4.Range(ws_4.Cells(1,5),ws_4.Cells(100000,50)).ClearContents()
        delete_all_records()        

class wsEvents_showoverride:
    def OnClick(self,*args):
        x = get_data()
        df = pd.DataFrame(x)
        li = ['Material_Number', 'Region', 'Delivering_Plant', 'Ship-To_Region',
       'Sold-To_Customerr_Name', 'Standard_price_USD', 'Lead_Time',
       'Service_Level', 'Forecast_Bucket', 'Forecast_Period', 'SafetyFactor',
       'safety_stock', 'ROP', 'Max_Stock', 'Churn_in_Dollar', 'Date',
       'Quantity_Delivered', 'safety_stock_override', 'ROP_override',
       'Max_Stock_override', 'Churn_in_Dollar_override', 'Override_date',
       'Predicted']
        df.columns = li
        df2_Var = df.columns[:15].to_list() + df.columns[18:-1].to_list()
        check_pivot = df.pivot_table(index=df2_Var,columns='Date',values='Quantity_Delivered').reset_index(drop=False)
        print(df['Quantity_Delivered'])
        StartRow = 2
        StartCol = 5
        ws_4.Range(ws_4.Cells(1,5),ws_4.Cells(100000,50)).ClearContents()
        ws_4.Range(ws_4.Cells(StartRow-1,StartCol),# Cell to start the "paste"
                ws_4.Cells(StartRow-1,StartCol+len(check_pivot.columns)-1)).Value = check_pivot.columns
        ws_4.Range(ws_4.Cells(StartRow,StartCol),# Cell to start the "paste"
                ws_4.Cells(StartRow+len(check_pivot)-1,
                        StartCol+len(check_pivot.columns)-1)# No -1 for the index
                ).Value = check_pivot.to_records(index=False)

class wsEvents_override:
    def OnClick(self,*args):
        # searches all rows and columns.
        df = pd.read_excel(os.path.join(os.path.dirname(__file__),"output.xlsx"))
        #df = pd.read_excel(r'D:\Python_Work\spd_analysis\output.xlsx')
        #df.drop(df.columns[-1],axis = 1,inplace = True)
        stock_column = df.columns[-5:].to_list()
        id_var = df.columns[:10].to_list() + stock_column
        print(id_var)
        # compare code

        ## changed df
        df_temp = pd.DataFrame(ws_2.Range(ws_2.Cells(8,6),# Cell to start the "paste"
                ws_2.Cells(8+len(df)-1,
                        6+len(df.columns)-1)# No -1 for the index
                ).Value)

        df_temp.columns = df.columns
        x = df.iloc[:,-int(df['Forecast_Period'].mean()+5):-5].compare(df_temp.iloc[:,-int(df_temp['Forecast_Period'].mean()+5):-5],align_axis=0)
        
        #concat the sheet
        df_forecast = pd.DataFrame()
        df_override = pd.DataFrame()
        for i in range(0,len(x.index),2):
            df_fVar = pd.DataFrame(df.iloc[x.index[i][0]]).transpose()
            df_oVar = pd.DataFrame(df_temp.iloc[x.index[i][0]]).transpose()
            df_forecast = pd.concat([df_forecast,df_fVar],axis=0)
            df_override = pd.concat([df_override,df_oVar],axis=0)    
        df_f = df_forecast.melt(id_vars=id_var,var_name='Date',value_name='Quantity_Delivered')
        df_override = df_override.drop(df_override.columns[-4:],axis=1).reset_index(drop=True)
        #calculate stock
        period = df_override['Forecast_Period'].mean()
        ss_list = safety_stock(df_override,CurrentValueVar,1)
        df_override['safety_stock_override']=[ round(elem) for elem in ss_list ]

        ## ROP Calculation:
        rop_list = ROP(df_override,CurrentValueVar,1)
        df_override['ROP_override'] = [ round(elem) for elem in rop_list ]

        #Maximum stock calculation:
        ms_list = max_stock(df_override,CurrentValueVar,1)
        df_override['Max_Stock_override'] = ms_list
        # Churn in Dollar 
        churn_list = churn_in_dollar(df_override,CurrentValueVar,1)
        df_override['Churn_in_Dollar_override'] = [ round(elem) for elem in churn_list ]

        stock_column_over = df_override.columns[-5:].to_list()
        id_var_over = df_override.columns[:10].to_list() + stock_column_over
        df_override = df_override.melt(id_vars=id_var_over,var_name='Date_override',value_name='Quantity_Delivered_override')
        for row in df_override.itertuples():
            if bool(re.match(r"[MmQq]", row.Date_override)) == True:
                df_override.Date_override[row.Index] = df_override.Date_override[row.Index] + '_Override'
        df_override_lastcolumns = df_override.iloc[:,-6:]
        df_final = pd.concat([df_f,df_override_lastcolumns],axis=1)
        #slicing dataset
        df_slice = df_final[df_final['Date_override'].str.endswith('Override')]
        df_check = pd.concat([df_final,df_slice],axis=0,ignore_index=True)
        df_check['Date'].iloc[-len(df_slice):] = df_check['Date_override'].iloc[-len(df_slice):]
        df_check['Quantity_Delivered'].iloc[-len(df_slice):] = df_check['Quantity_Delivered_override'].iloc[-len(df_slice):]
        df_final = df_check.drop(df_check.columns[-2:],axis = 1)
        df_final['Override_date'] = datetime.datetime.now()
        df_final['Override_date'] = df_final['Override_date'].apply(lambda x: x.strftime('%m/%d/%Y, %H:%M:%S'))
        df_final['Predicted'] = 0
        for row in df_final.itertuples():
            if bool(re.match(r"[MmQq]", row.Date)) == True:
                df_final.Predicted[row.Index] = 1
        df_final['Delivering_Plant'] = df_final['Delivering_Plant'].astype(str)      
        df_final[['Standard_price_USD', 'Lead_Time','Service_Level','Forecast_Bucket','Forecast_Period','SafetyFactor','safety_stock'
                        ,'ROP','Max_Stock','Churn_in_Dollar','Quantity_Delivered','Churn_in_Dollar_override']] = df_final[['Standard_price_USD', 'Lead_Time','Service_Level',
                            'Forecast_Bucket','Forecast_Period','SafetyFactor','safety_stock','ROP','Max_Stock',
                                'Churn_in_Dollar','Quantity_Delivered','Churn_in_Dollar_override']].apply(pd.to_numeric)
        print(df_final.dtypes)                        
        li_1 = df_final.to_numpy().tolist()
        #delete_table()
        check_table()
        check_table_column()
        insert_to_override(li_1)

class WsEvents_MLmodel:
    def OnClick(self,*args):
        print(*args,"---args----")
        workBook.WorkSheets(2).Activate()
        CurrentValueVar = readCurrentValue(ws_1)
        Model_output = ML_Model_Stock_calc(data1,CurrentValueVar)
        # Transform for database
        #result_df['Override'] = None
        result_df = Model_output[2]
        result_df['Service_Level'] = CurrentValueVar[1]
        result_df['Forecast_Bucket'] = CurrentValueVar[2]
        result_df['Forecast_Period'] = CurrentValueVar[5]
        df_temp_1 = result_df.iloc[:,-3:]
        df_temp_2 = result_df.iloc[:,:7]
        df_temp_3 = result_df.iloc[:,7:-3]
        part_info_temp = pd.concat([df_temp_2,df_temp_1],axis=1)
        result_df = pd.concat([part_info_temp,df_temp_3],axis = 1)
        li = np.array((result_df.columns).astype(str))
        result_df = result_df.reset_index(drop=True)
        result_df.to_excel(os.path.join(os.path.dirname(__file__),"output.xlsx"),index=False)

        #highlights
        try:
            x = get_data()
            df_over = pd.DataFrame(x)
            li1 = ['Material_Number', 'Region', 'Delivering_Plant', 'Ship-To_Region',
            'Sold-To_Customerr_Name', 'Standard_price_USD', 'Lead_Time',
            'Service_Level', 'Forecast_Bucket', 'Forecast_Period', 'SafetyFactor',
            'safety_stock', 'ROP', 'Max_Stock', 'Churn_in_Dollar', 'Date',
            'Quantity_Delivered', 'safety_stock_override', 'ROP_override',
            'Max_Stock_override', 'Churn_in_Dollar_override', 'Override_date',
            'Predicted']
            df_over.columns = li1
            df_over = df_over.iloc[:,0:10].drop_duplicates()
            df_for = result_df.iloc[:,0:10]
            df_for['index']= df_for.index
            matchVar = pd.merge(df_for,df_over,how='inner')
        except:
            pass

        ws_outputsheet = workBook.Worksheets("Forecast")
        ws_outputsheet.Cells(3,13).Value = len(data1)
        ws_outputsheet.Cells(4,13).Value = Model_output[0]*100
        ws_outputsheet.Cells(5,13).Value = Model_output[1]*100

        StartRow = 8
        StartCol = 6
        #ws_outputsheet.Range(ws_outputsheet.Cells(StartRow-1,6),ws_outputsheet.Cells(100000,60)).Font.ColorIndex = 0
        ws_outputsheet.Range(ws_outputsheet.Cells(StartRow,6),ws_outputsheet.Cells(100000,15)).Interior.Color = rgbToInt([255,255,255])

        try:
            for i in matchVar['index']:
                ws_outputsheet.Range(ws_outputsheet.Cells(8+i,6),ws_outputsheet.Cells(8+i,15)).Interior.Color = rgbToInt((100,255,200)) # green
        except:
            pass
        ws_outputsheet.Range(ws_outputsheet.Cells(StartRow-1,6),ws_outputsheet.Cells(100000,60)).ClearContents()
        ws_outputsheet.Range(ws_outputsheet.Cells(StartRow-1,StartCol),# Cell to start the "paste"
                ws_outputsheet.Cells(StartRow-1,StartCol+len(result_df.columns)-1)).Value = li
        ws_outputsheet.Range(ws_outputsheet.Cells(StartRow,StartCol),# Cell to start the "paste"
                ws_outputsheet.Cells(StartRow+len(result_df)-1,
                        StartCol+len(result_df.columns)-1)# No -1 for the index
                ).Value = result_df.to_records(index=False)
        


if __name__ == '__main__':
    #defining win32com object
    excel= win32.dynamic.Dispatch('Excel.Application')
    workBook = excel.Workbooks('UI_Tool.xlsm')
    ws_1 = workBook.Worksheets(1)
    ws_3 = workBook.Worksheets(3)
    ws_2 = workBook.Worksheets(2)
    ws_4 = workBook.Worksheets(4)


    #Data Cleaning: Whenever app is open
    data1 = pd.read_excel(os.path.join(os.path.dirname(__file__),"Demand_Planning_Report_new.xlsx"))
    #data1 = pd.read_excel(r'D:\Python_Work\spd_analysis\Demand_Planning_Report.xlsx')
    data1 = data1.drop_duplicates()
    data1.columns = data1.columns.str.replace(' ','_')
    data1['Actual_Goods_Movement_Date'] = pd.to_datetime(data1['Actual_Goods_Movement_Date'])
    data1['Line_Creation_Date'] = pd.to_datetime(data1['Line_Creation_Date'])
    data1.columns.values[7] = 'Quantity_Delivered_Actual'
    data1.columns.values[29] = 'Lead_Time'
    data1['Ship-To_Region'] = data1.apply(lambda x: 'Unknown' if len(x['Ship-To_Region'])<=1 else x['Ship-To_Region']  ,axis=1)


    df_demand = data1.copy()
    America_plant = [10,12,1900,2900,8003,8047,8012]
    APAC_plant = [16,5200,5210,5250,5400,6900,6910,8004,8005,8006,8011,8084,8085,5100]
    Europe_plant = [4350]

    df_demand['Region'] = None
    df_demand.loc[df_demand['Delivering_Plant'].isin(America_plant),'Region'] = 'America'
    df_demand.loc[df_demand['Delivering_Plant'].isin(APAC_plant),'Region'] = 'APAC'
    df_demand.loc[df_demand['Delivering_Plant'].isin(Europe_plant),'Region'] = 'Europe'


    df_demand.to_excel(os.path.join(os.path.dirname(__file__),"clean_demand_data.xlsx"),index=False)

    df_demand = pd.read_excel(os.path.join(os.path.dirname(__file__),"clean_demand_data.xlsx"))
    data1 = df_demand.copy()
    print(data1,"----data1----")
    print(type(data1),"---type-----")

    ## Working with data to find unique customer, region, and Plant/Depot
    #unique_list
    unique_part,unique_plant_depot,unique_customers,unique_region,unique_ship_region = np.unique(df_demand['Material_Number']),np.unique(df_demand['Delivering_Plant']), np.unique(df_demand['Sold-To_Customerr_Name']),np.unique(df_demand['Region']),np.unique(df_demand['Ship-To_Region'])
    unique_part,unique_plant_depot,unique_customers,unique_region,unique_ship_region = np.vstack(unique_part),np.vstack(unique_plant_depot),np.vstack(unique_customers),np.vstack(unique_region),np.vstack(unique_ship_region)

    ### Filter Value for list

    ws_1.Range(ws_1.Cells(7,6),ws_1.Cells(7 + 10000,6)).ClearContents()
    ws_1.Range(ws_1.Cells(7,6),ws_1.Cells(7 + len(unique_part)-1,6)).Value = unique_part
    ws_1.Range(ws_1.Cells(7,8),ws_1.Cells(7 + 10000,8)).ClearContents()
    ws_1.Range(ws_1.Cells(7,8),ws_1.Cells(7 + len(unique_plant_depot)-1,8)).Value = unique_plant_depot
    ws_1.Range(ws_1.Cells(7,10),ws_1.Cells(7 + 10000,10)).ClearContents()
    ws_1.Range(ws_1.Cells(7,10),ws_1.Cells(7 + len(unique_customers)-1,10)).Value = unique_customers
    ws_1.Range(ws_1.Cells(7,7),ws_1.Cells(7 + 10000,7)).ClearContents()
    ws_1.Range(ws_1.Cells(7,7),ws_1.Cells(7 + len(unique_region)-1,7)).Value = unique_region
    ws_1.Range(ws_1.Cells(7,9),ws_1.Cells(7 + 10000,9)).ClearContents()
    ws_1.Range(ws_1.Cells(7,9),ws_1.Cells(7 + len(unique_ship_region)-1,9)).Value = unique_ship_region
    ws_1.Range(ws_1.Cells(7,11),ws_1.Cells(7 + 10000,11)).ClearContents()
    defaultValue(6,14,ws_1)


    CurrentValueVar = readCurrentValue(ws_1)
    # print(CurrentValueVar,"-----current value----")
    # print(type(CurrentValueVar),"---type----")
    xl_events=win32.WithEvents(ws_1.OLEObjects("CommandButton3").Object,wsEvents_filterPN)
    xl_events_1 = win32.WithEvents(ws_1.OLEObjects("CommandButton4").Object,wsEvents_filterPlantDepo)
    xl_events_2 = win32.WithEvents(ws_1.OLEObjects("CommandButton5").Object,wsEvents_filterregion)
    xl_events_3 = win32.WithEvents(ws_1.OLEObjects("CommandButton6").Object,wsEvents_filtercus)
    xl_events_4 = win32.WithEvents(ws_1.OLEObjects("CommandButton7").Object,WsEvents_MLmodel)
    xl_events_5 = win32.WithEvents(ws_2.OLEObjects("CommandButton1").Object,wsEvents_override)
    xl_events_6 = win32.WithEvents(ws_1.OLEObjects("CommandButton1").Object,wsEvents_filtershipregion)
    #xl_events_7 = win32.WithEvents(ws_1.OLEObjects("CommandButton2").Object,WsEvents_GlobalDemand)
    xl_events_8 = win32.WithEvents(ws_4.OLEObjects("CommandButton1").Object,wsEvents_showoverride)
    xl_events_9 = win32.WithEvents(ws_4.OLEObjects("CommandButton2").Object,wsEvents_deleterecords)
    #xl_event_filterPlantDepo = win32.WithEvents(ws_1.OLEObjects("CommandButton4").Object,wsEvents)

    # define initalizer
    keepOpen = True
    while keepOpen:
        # display the message
        pythoncom.PumpWaitingMessages()


