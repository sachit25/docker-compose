#import packages
from unittest import result
import scipy.stats as st
from scipy.stats import norm
from math import sqrt
import pandas as pd
import numpy as np
from function_file import *
import win32com.client as win32
from sklearn.tree import DecisionTreeRegressor
import sys
import os

def safety_factor(sample_len,s_level):
    if sample_len >= 30:
        return norm.ppf(s_level)  #  inverse of the CDF of the standard normal distribution
    elif sample_len < 30:
        return st.t.ppf(s_level , sample_len)  # inverse of the CDF of the T-Distribution   

def safety_stock_1(safety_factor,lead_time,demand_deviation,Period):
    return safety_factor*sqrt(lead_time/Period)*demand_deviation

def ROP(safety_stock, lead_time_demand):
    return safety_stock + lead_time_demand 
    
def Max_stock(predicted_demand,safety_stock):
    return predicted_demand + safety_stock  

def safFactor_Location(data,serviceLevel):
    safety_factor_list = []
    partnumber_list =[]
    region_list =[]
    plant_list = []
    ship_region_list=[]
    customer_list = []
    for i in data['Material_Number'].unique():
        df_material = data[data['Material_Number']==i]
        for y in df_material['Region'].unique():
            df_region = df_material[df_material['Region']==y]
            for z in df_region['Delivering_Plant'].unique():
                df_plantDepo = df_region[df_region['Delivering_Plant']==z]
                for n in df_plantDepo['Ship-To_Region'].unique():
                    df_ship_region = df_plantDepo[df_plantDepo['Ship-To_Region']==n]
                    for m in df_ship_region['Sold-To_Customerr_Name'].unique():
                        df_customer = df_ship_region[df_ship_region['Sold-To_Customerr_Name']==m]
                        servLevelVar = (serviceLevel/100)
                        lengthData = int(len(df_customer))
                        safFactor = safety_factor(lengthData,servLevelVar)
                        safety_factor_list.append(safFactor)
                        partnumber_list.append(i)
                        region_list.append(y)
                        plant_list.append(z)
                        ship_region_list.append(n)
                        customer_list.append(m)

    return safety_factor_list,partnumber_list,region_list,plant_list,ship_region_list,customer_list

def datasets(df, lookIntoPeriod, OutputNextPeriod, y_test_len):
    D = df.values
    periods = D.shape[1]
    # Training set creation: run through all the possible time windows
    loops = periods + 1 - lookIntoPeriod - OutputNextPeriod - y_test_len

    train = []
    for col in range(1,loops):
        list_1= D[:,col:col+lookIntoPeriod+OutputNextPeriod]
        train.append(list_1)

    train = np.vstack(train)
    X_train, Y_train = np.split(train,[lookIntoPeriod],axis=1)

    # Test set creation: unseen “future” data with the demand just before
    max_col_test = periods - lookIntoPeriod - OutputNextPeriod + 1
    test = []
    for col in range(loops,max_col_test):
        list_1= D[:,col:col+lookIntoPeriod+OutputNextPeriod]
        test.append(list_1)

    test = np.vstack(test)
    X_test, Y_test = np.split(test,[lookIntoPeriod],axis=1)

    # this data formatting is needed if we only predict a single period
    if OutputNextPeriod == 1:
        Y_train = Y_train.ravel()
        Y_test = Y_test.ravel()
    
    return X_train, Y_train, X_test, Y_test

def safety_stock(data,CurrentValue,PeriodVar):
    CurrentValueVar = CurrentValue
    if PeriodVar == 0:
        forecast_period = CurrentValueVar[5]
    else:
        forecast_period = data['Forecast_Period'].mean()
    ss_list = []
    for row in data.itertuples():
        leadTimVar = row.Lead_Time
        if leadTimVar == 0:
            ss_list.append(0)
            continue
        SafFacVar = row.SafetyFactor
        if CurrentValueVar[0] =='Monthly':
            Forperiod = forecast_period*30
        elif CurrentValueVar[0] == 'Quarterly':
            Forperiod = forecast_period*90

        period_list = list(row[-(int(forecast_period)+1):-1])
        period_list = [x+0.01 for x in period_list]
        StdDev = np.std(period_list)
        ss = safety_stock_1(SafFacVar,leadTimVar,StdDev,Forperiod)
        ss_list.append(ss)
    return ss_list

def ROP(data,CurrentValue,PeriodVar):
    CurrentValueVar = CurrentValue
    if PeriodVar == 0:
        forecast_period = CurrentValueVar[5]
    else:
        forecast_period = data['Forecast_Period'].mean()
    rop_list = []
    for row in data.itertuples():
        TotalDemand = sum(list(row[-(int(forecast_period)+2):-2]))
        if CurrentValueVar[0] =='Monthly':
            Forperiod = forecast_period*30
        elif CurrentValueVar[0] == 'Quarterly':
            Forperiod = forecast_period*90

        LeadperiodRatio = row.Lead_Time/Forperiod
        LeadTimeDemand = TotalDemand * LeadperiodRatio
        if PeriodVar == 0:
            ss = row.safety_stock
        else:
            ss = row.safety_stock_override    
        rop_list.append(LeadTimeDemand + ss)            
    return rop_list

def max_stock(data,CurrentValue,PeriodVar):
    CurrentValueVar = CurrentValue
    if PeriodVar == 0:
        forecast_period = CurrentValueVar[5]
    else:
        forecast_period = data['Forecast_Period'].mean()
    ms_list = []
    for row in data.itertuples():
        TotalDemand = round(sum(list(row[-(int(forecast_period)+3):-3])))
        if PeriodVar ==0:
            ss = row.safety_stock
        else:
            ss = row.safety_stock_override    
        ms = TotalDemand + ss
        ms_list.append(ms)
    return ms_list

def churn_in_dollar(data,CurrentValue,PeriodVar):
    CurrentValueVar  =CurrentValue
    if PeriodVar == 0:
        forecast_period = CurrentValueVar[5]
    else:
        forecast_period = data['Forecast_Period'].mean()
    churn_list = []
    for row in data.itertuples():
        TotalDemand = round(sum(list(row[-(int(forecast_period)+4):-4])))
        sp = row.Standard_price_USD
        churn = TotalDemand * sp
        churn_list.append(churn)
    return churn_list     
   
def ML_Model_Stock_calc(data,CurrentValue):
    CurrentValueVar = CurrentValue
    df_demand = data.iloc[:,[0,2,7,8,9,10,17,18,29,33,34]]
    Period = CurrentValueVar[0]
    df_periods_quantity = period_wise_grouping(data,df_demand,Period,pd)
    df_plant = plant_wise_grouping(data,df_demand,Period,pd)
    df_concat = pd.concat([df_plant,df_periods_quantity],axis = 0)

    # # Preparing data for ML model
    df_transform = pd.pivot_table(data=df_concat, values='Quantity_Delivered_Actual', index=['Material_Number','Region','Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name'], columns='Period', fill_value=0)
    df_df = df_transform.copy()
    df_df.reset_index(inplace=True)
    ForecastBucket = int(CurrentValueVar[2])
    ForecastPeriod = int(CurrentValueVar[5])
    X_train, Y_train, X_test, Y_test = datasets(df_transform,ForecastBucket,ForecastPeriod,ForecastPeriod)

    # — Instantiate a Decision Tree Regressor 
    tree = DecisionTreeRegressor(max_depth=5,min_samples_leaf=1) 
    # — Fit the tree to the training data 
    tree.fit(X_train,Y_train)
    # Create a prediction based on our model 
    Y_train_pred = tree.predict(X_train) 
    # Compute the Mean Absolute Error of the model
    MAE_tree = np.mean(abs(Y_train - Y_train_pred))/(np.mean(Y_train) + .01)
    # Print the results 
    print('Tree on train set MAE%:',round(MAE_tree*100,1))
    Y_test_pred = tree.predict(X_test)
    MAE_test = np.mean(abs(Y_test - Y_test_pred))/(np.mean(Y_test) + 0.01)
    print('Tree on test set MAE%:',round(MAE_test*100,1))

    # # Making prediction set

    #filter dataset based on input values:
    partNumber = np.vstack(CurrentValueVar[6])
    partNumber = partNumber[partNumber != np.array(None)]
    if len(partNumber)!=0:
        df_demand = df_demand[df_demand["Material_Number"].isin(partNumber)]

    regionList = np.vstack(CurrentValueVar[9])
    regionList = regionList[regionList != np.array(None)]
    if len(regionList)!=0:
        df_demand = df_demand[df_demand["Region"].isin(regionList)]
        
    plantList = np.vstack(CurrentValueVar[7])
    plantList = plantList[plantList != np.array(None)]
    if len(plantList)!=0:
        df_demand = df_demand[df_demand["Delivering_Plant"].isin(plantList)]
    shipregionList = np.vstack(CurrentValueVar[11])
    shipregionList = shipregionList[shipregionList != np.array(None)]
    if len(shipregionList)!=0:
        df_demand = df_demand[df_demand["Ship-To_Region"].isin(shipregionList)]

    cusList = np.vstack(CurrentValueVar[8])
    cusList = cusList[cusList != np.array(None)]
    if len(cusList)!=0:
        df_demand = df_demand[df_demand["Sold-To_Customerr_Name"].isin(cusList)]

    if CurrentValueVar[3] != None:
        df_demand = df_demand[df_demand['Standard_price_USD'] > CurrentValueVar[3]]
    if CurrentValueVar[4] != None:
        df_demand = df_demand[df_demand['Standard_price_USD'] > CurrentValueVar[4]]

    if df_demand.empty:
        print('data not adequate')
        sys.exit(1)

    # Customer-wise consolidation:
    consolidate_list = np.vstack(CurrentValueVar[10])
    consolidate_list = consolidate_list[consolidate_list != np.array(None)]

    if len(cusList) ==0 and len(shipregionList) == 0 and len(regionList) != 0 and len(plantList) !=0:
        df_test = plant_wise_grouping(data,df_demand,Period,pd)


    elif len(cusList) ==0 and len(shipregionList) == 0 and len(regionList) == 0 and len(plantList) == 0:
        df_test = part_wise_grouping(data,df_demand,Period,pd)    
    
    elif len(cusList) ==0 and len(shipregionList) == 0  and len(plantList) == 0 and len(regionList) !=0:
        df_test = global_grouping(data,df_demand,Period,pd)

    elif len(cusList)==0 and len(shipregionList) !=0 and len(plantList) !=0 and len(regionList)!=0:
            df_grouped = pd.DataFrame()
            Period = 'Monthly'
            input_partlist = df_demand['Material_Number'].unique()
            for i in input_partlist:
                df_material = df_demand[df_demand['Material_Number']==i]
                for y in df_material['Delivering_Plant'].unique():
                    df_plant = df_material[df_material['Delivering_Plant']==y]
                    for z in df_plant['Ship-To_Region'].unique():
                        df_shipregion = df_plant[df_plant['Ship-To_Region']==z]
                        df_copy = df_shipregion.copy()
                        df_copy = df_copy.drop_duplicates()
                        df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                        if Period == 'Monthly':
                            df_copy['Actual_Goods_Movement_Date'] = pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                            idx = pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_demand['Actual_Goods_Movement_Date'].max(), freq='M')
                        elif Period == 'Quarterly':
                            df_copy['Actual_Goods_Movement_Date'] = pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                            idx =pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Q')   
                        elif Period == 'Yearly':
                            df_copy['Actual_Goods_Movement_Date'] = pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                            idx = pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Y')
                        elif Period =='Half Yearly':
                            df_copy['Actual_Goods_Movement_Date'] =pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                            idx = pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='6M')
                        df_copy =df_copy.iloc[:,[2,5]]
                        df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                        df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                        df_copy.reset_index(drop=False,inplace=True)
                        df_copy['Material_Number'] = i
                        df_copy['Region'] = df_plant['Region'].unique()[0]
                        df_copy['Delivering_Plant'] = y
                        df_copy['Ship-To_Region'] = z
                        df_copy['Sold-To_Customerr_Name'] = 'consolidate'
                        df_copy.columns.values[0] = 'Period'
                        df_grouped = pd.concat([df_grouped,df_copy],axis=0)

            df_test = df_grouped            
    else:
        if len(consolidate_list)!=0:
            df_demand_list1 = df_demand[df_demand['Sold-To_Customerr_Name'].isin(consolidate_list)]
            df_demand_list2 = df_demand[~df_demand['Sold-To_Customerr_Name'].isin(consolidate_list)]  
            #df_test1 = plant_wise_grouping(data1,df_demand_list1,Period,pd)
            df_grouped = pd.DataFrame()
            Period = 'Monthly'
            input_partlist = df_demand_list1['Material_Number'].unique()
            for i in input_partlist:
                df_material = df_demand_list1[df_demand_list1['Material_Number']==i]
                for y in df_material['Delivering_Plant'].unique():
                    df_plant = df_material[df_material['Delivering_Plant']==y]
                    for z in df_plant['Ship-To_Region'].unique():
                        df_shipregion = df_plant[df_plant['Ship-To_Region']==z]
                        if df_shipregion['Sold-To_Customerr_Name'].nunique()>1:
                            df_copy = df_shipregion.copy()
                            df_copy = df_copy.drop_duplicates()
                            df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                            if Period == 'Monthly':
                                df_copy['Actual_Goods_Movement_Date'] = pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                                idx = pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_demand['Actual_Goods_Movement_Date'].max(), freq='M')
                            elif Period == 'Quarterly':
                                df_copy['Actual_Goods_Movement_Date'] = pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                                idx =pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Q')   
                            elif Period == 'Yearly':
                                df_copy['Actual_Goods_Movement_Date'] = pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                                idx = pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Y')
                            elif Period =='Half Yearly':
                                df_copy['Actual_Goods_Movement_Date'] =pd.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                                idx = pd.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='6M')
                            df_copy =df_copy.iloc[:,[2,5]]
                            df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                            df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                            df_copy.reset_index(drop=False,inplace=True)
                            df_copy['Material_Number'] = i
                            df_copy['Region'] = df_plant['Region'].unique()[0]
                            df_copy['Delivering_Plant'] = y
                            df_copy['Ship-To_Region'] = z
                            df_copy['Sold-To_Customerr_Name'] = 'consolidate'
                            df_copy.columns.values[0] = 'Period'
                            df_grouped = pd.concat([df_grouped,df_copy],axis=0)
                
            df_test1 = df_grouped
            df_test2 = period_wise_grouping(data,df_demand_list2,Period,pd)
            df_test = pd.concat([df_test1,df_test2],axis=0)
        else:
            df_test = period_wise_grouping(data,df_demand,Period,pd)     

    df_test_transform = pd.pivot_table(data=df_test, values='Quantity_Delivered_Actual', index=['Material_Number','Region','Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name'], columns='Period', fill_value=0) 
    df_test_transform.reset_index(inplace=True)
    print(df_test_transform.iloc[:,:6])
    df_demand = df_demand.iloc[:,[0,1,6,7,8,9,10]]

    if len(cusList)==0 and len(shipregionList) == 0 and len(regionList)!=0 and len(plantList)!=0:
            df_demand_sp = df_demand.groupby(['Material_Number','Region','Delivering_Plant']).Standard_price_USD.mean().reset_index(drop=False)
            df_demand_sp[['Ship-To_Region','Sold-To_Customerr_Name']] = 'consolidate'
            df_demand_lt = df_demand.groupby(['Material_Number','Region','Delivering_Plant']).Lead_Time.mean().reset_index(drop=False)
            df_demand_lt[['Ship-To_Region','Sold-To_Customerr_Name']] = 'consolidate'
            df_test_transform =pd.merge(df_test_transform,df_demand_sp,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
            df_test_transform =pd.merge(df_test_transform,df_demand_lt,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
    elif len(cusList) == 0 and len(shipregionList)!=0 and len(regionList)!=0 and len(plantList)!=0:
            df_demand_sp = df_demand.groupby(['Material_Number','Region','Delivering_Plant','Ship-To_Region']).Standard_price_USD.mean().reset_index(drop=False)
            df_demand_sp['Sold-To_Customerr_Name'] = 'consolidate'
            df_demand_lt = df_demand.groupby(['Material_Number','Region','Delivering_Plant','Ship-To_Region']).Lead_Time.mean().reset_index(drop=False)
            df_demand_lt['Sold-To_Customerr_Name'] = 'consolidate'
            df_test_transform =pd.merge(df_test_transform,df_demand_sp,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
            df_test_transform =pd.merge(df_test_transform,df_demand_lt,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
    elif len(cusList)==0 and len(shipregionList) == 0 and len(regionList)==0 and len(plantList)==0:
            df_demand_sp = df_demand.groupby(['Material_Number']).Standard_price_USD.mean().reset_index(drop=False)
            df_demand_sp[['Region','Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name']] = 'consolidate'
            df_demand_lt = df_demand.groupby(['Material_Number']).Lead_Time.mean().reset_index(drop=False)
            df_demand_lt[['Region','Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name']] = 'consolidate'
            df_test_transform =pd.merge(df_test_transform,df_demand_sp,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
            df_test_transform =pd.merge(df_test_transform,df_demand_lt,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
    elif len(cusList) ==0 and len(shipregionList) == 0  and len(plantList) == 0 and len(regionList) !=0:
            df_test_transform = pd.pivot_table(data=df_test, values='Quantity_Delivered_Actual', index=['Material_Number','Region','Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name'], columns='Period', fill_value=0) 
            df_test_transform.reset_index(inplace=True)
            df_demand_sp = df_demand.groupby(['Material_Number','Region']).Standard_price_USD.mean().reset_index(drop=False)
            df_demand_sp[['Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name']] = 'consolidate'
            df_demand_lt = df_demand.groupby(['Material_Number','Region']).Lead_Time.mean().reset_index(drop=False)
            df_demand_lt[['Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name']] = 'consolidate'
            df_test_transform =pd.merge(df_test_transform,df_demand_sp,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
            df_test_transform =pd.merge(df_test_transform,df_demand_lt,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
    else:
        if len(consolidate_list)!=0:
            df_demand_list1 = df_demand_list1.iloc[:,[0,1,6,7,8,9,10]]
            df_demand_list2 = df_demand_list2.iloc[:,[0,1,6,7,8,9,10]]
            df_demand_sp = df_demand_list1.groupby(['Material_Number','Region','Delivering_Plant','Ship-To_Region']).Standard_price_USD.mean().reset_index(drop=False)
            df_demand_lt = df_demand_list1.groupby(['Material_Number','Region','Delivering_Plant','Ship-To_Region']).Lead_Time.mean().reset_index(drop = False)
            df_demand_lt['Sold-To_Customerr_Name'] = 'consolidate'
            df_demand_sp['Sold-To_Customerr_Name'] = 'consolidate'
            df_test_transform_cons = df_test_transform[df_test_transform['Sold-To_Customerr_Name']=='consolidate']
            df_test_transform_notcons = df_test_transform[df_test_transform['Sold-To_Customerr_Name']!='consolidate']
            df_test_transform_cons =pd.merge(df_test_transform_cons,df_demand_sp,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
            df_test_transform_cons =pd.merge(df_test_transform_cons,df_demand_lt,on=['Material_Number','Delivering_Plant','Region','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
            df_test_transform_notcons =pd.merge(df_test_transform_notcons,df_demand_list2,on=['Material_Number','Delivering_Plant','Sold-To_Customerr_Name','Region','Ship-To_Region'],how='left')
            df_test_transform = pd.concat([df_test_transform_cons,df_test_transform_notcons],axis=0)
            print(df_test_transform_cons.iloc[:,:6])
            print(df_test_transform_notcons.iloc[:,:6])   
        else:
            df_test_transform =pd.merge(df_test_transform,df_demand,on=['Material_Number','Delivering_Plant','Sold-To_Customerr_Name','Region','Ship-To_Region'],how='left')


    df_test_transform.drop_duplicates(inplace=True)
    df_test_transform.reset_index(drop=True,inplace = True)
    part_info_1 = pd.DataFrame(df_test_transform.iloc[:,0:5])
    part_info_2 = pd.DataFrame(df_test_transform.iloc[:,-2:])
    part_info = pd.concat([part_info_1,part_info_2],axis=1)
    data_fields_test = df_test_transform.iloc[:,-(ForecastBucket+2):-2]
    test_df = pd.concat([part_info,data_fields_test],axis=1)
    test_set = test_df.iloc[:,-ForecastBucket:].values
    try:
        prediction = pd.DataFrame(tree.predict(test_set))
        prediction = pd.DataFrame(prediction.values.round())
        if Period == 'Monthly':
            prediction = prediction.add_prefix('Month_')
        if Period == 'Quaterly':
            prediction = prediction.add_prefix('Quarterly_')
    except:
        print("Data is not adequate for prediction")  
        sys.exit(1)

    list_sf = safFactor_Location(df_test,CurrentValueVar[1])
    safety_factor_list = {'Material_Number':list_sf[1],'Region':list_sf[2],'Delivering_Plant':list_sf[3],'Ship-To_Region':list_sf[4],'Sold-To_Customerr_Name':list_sf[5],'SafetyFactor':list_sf[0]}
    safety_factor_df = pd.DataFrame(safety_factor_list)
    result_df = pd.concat([test_df,prediction],axis=1)
    result_df = pd.merge(result_df,safety_factor_df,on=['Material_Number','Region','Delivering_Plant','Ship-To_Region','Sold-To_Customerr_Name'],how='left')
    ss_list = safety_stock(result_df,CurrentValueVar,0)
    result_df['safety_stock']=[ round(elem) for elem in ss_list ]
    rop_list = ROP(result_df,CurrentValueVar,0)
    result_df['ROP'] = [ round(elem) for elem in rop_list ]
    ms_list = max_stock(result_df,CurrentValueVar,0)
    result_df['Max_Stock'] = ms_list
    churn_list = churn_in_dollar(result_df,CurrentValueVar,0)
    result_df['Churn_in_Dollar'] = [ round(elem) for elem in churn_list ]
    
    output_list = [MAE_tree,MAE_test,result_df]
    return output_list


