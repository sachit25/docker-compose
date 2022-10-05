from distutils.sysconfig import customize_compiler
# from h11 import Data
from datetime import date
import sqlite3

filter_plant = False
filter_region = False
filter_shipregion = False
filter_customer = False
filter_partno = False

#Default Values
def defaultValue(Startrow,Startcolumn,activeWorksheet):
    activeWorksheet.Cells(Startrow,Startcolumn).Value = 'Monthly'
    activeWorksheet.Cells(Startrow + 4,Startcolumn).Value = None
    activeWorksheet.Cells(Startrow + 5,Startcolumn).Value = 6
    activeWorksheet.Cells(Startrow + 6,Startcolumn).Value = None
    activeWorksheet.Cells(Startrow + 7,Startcolumn).Value = None
    activeWorksheet.Cells(Startrow + 8,Startcolumn).Value = None
    
# read current value:
def readCurrentValue(activeWorksheet):
    ListCurrentValue = []
    plant_list = activeWorksheet.Range("H7:H10000").Value
    partNumber =activeWorksheet.Range("F7:F10000").Value
    region_list = activeWorksheet.Range("G7:G10000").Value
    customer_list = activeWorksheet.Range("J7:J10000").Value
    DemandPeriodVar = activeWorksheet.Cells(6,14).Value
    ServiceLevelVar = activeWorksheet.Cells(10,14).Value
    ForecastBucket = activeWorksheet.Cells(11,14).Value
    Cost_LP = activeWorksheet.Cells(12,14).Value
    Cost_UP = activeWorksheet.Cells(13,14).Value
    ForecastPeriod = activeWorksheet.Cells(14,14).Value
    Consolidate_list = activeWorksheet.Range("K7:K1000").Value
    ShipToRegion_list = activeWorksheet.Range("I7:I10000").Value 
    ListCurrentValue.extend((DemandPeriodVar,
                                ServiceLevelVar,ForecastBucket,Cost_LP,Cost_UP,ForecastPeriod,partNumber,plant_list,customer_list,region_list,Consolidate_list,ShipToRegion_list))
    return ListCurrentValue

    
def InputChangeValidate(activeWorksheet_3,activeWorksheet_1,data,numpy_object,field,fieldValueList,changeToOtherField):
    partVar = numpy_object.vstack(fieldValueList[6])
    partVar = partVar[partVar != numpy_object.array(None)]
    regionVar = numpy_object.vstack(fieldValueList[9])
    regionVar = regionVar[regionVar != numpy_object.array(None)]
    plantVar = numpy_object.vstack(fieldValueList[7])
    plantVar = plantVar[plantVar != numpy_object.array(None)]
    cusVar = numpy_object.vstack(fieldValueList[8])
    cusVar = cusVar[cusVar != numpy_object.array(None)]
    shipregionVar = numpy_object.vstack(fieldValueList[11])
    shipregionVar = shipregionVar[shipregionVar != numpy_object.array(None)]
    global filter_customer,filter_plant,filter_region,filter_shipregion,filter_partno
    if field == 'PartNumber':
        dfFilter = data
        if len(partVar) != 0:
            dfFilter = data[data["Material_Number"].isin(partVar)]
        else:
            filter_partno = False
            dfFilter = data
        filter_partno = True        
        filter_customer = False
        filter_plant = False
        filter_region = False
        filter_shipregion = False
        InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Region',fieldValueList,1)
        #InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'DepoPlant',fieldValueList,1)
        #InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Customer',fieldValueList,1)
        #InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Ship_Region',fieldValueList,1)

    
    if field =='Region':
        if changeToOtherField == 0:
            filter_region = True
            dfFilter = data
            if len(partVar)!=0 and filter_partno:
                dfFilter = data[data['Material_Number'].isin(partVar)]
                if len(regionVar)!=0:
                    dfFilter = dfFilter[dfFilter["Region"].isin(regionVar)]
                else:
                    filter_region = False    
            else:
                if len(regionVar) !=0:
                    dfFilter = data[data["Region"].isin(regionVar)]
                else:
                    filter_region = False
                    dfFilter = data            
        else:
            dfFilter = data

        uniqueregion = numpy_object.unique(dfFilter['Region'])
        uniqueregion = numpy_object.vstack(uniqueregion)
        uniquePartnumber = numpy_object.unique(dfFilter['Material_Number'])
        uniquePartnumber = numpy_object.vstack(uniquePartnumber)
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + 10000,7)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + len(uniqueregion)-1,7)).Value = uniqueregion
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + 10000,6)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + len(uniquePartnumber)-1,6)).Value = uniquePartnumber
        activeWorksheet_3.Range(activeWorksheet_3.Cells(4,22),activeWorksheet_3.Cells(4 + 60000,22)).ClearContents()
        if changeToOtherField == 0:
            InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'DepoPlant',fieldValueList,1)
        else:
            InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'DepoPlant',fieldValueList,1)   

            
    #unique Depot/PLant
    if field == 'DepoPlant':
        if changeToOtherField == 0:
            dfFilter = data
            filter_plant = True
            if len(partVar)!=0 and filter_partno:
                dfFilter = dfFilter[dfFilter['Material_Number'].isin(partVar)]
                if len(regionVar)!=0 and filter_region:
                    dfFilter = dfFilter[dfFilter['Region'].isin(regionVar)]
                    if len(shipregionVar)!=0 and filter_shipregion:
                        dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                    else:
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False       

                else:
                    if len(shipregionVar)!=0 and filter_shipregion:
                        dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                    else:
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False           
            else:
                if len(regionVar)!=0 and filter_region:
                    dfFilter = dfFilter[dfFilter['Region'].isin(regionVar)]
                    if len(shipregionVar)!=0 and filter_shipregion:
                        dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                    else:
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False       

                else:
                    if len(shipregionVar)!=0 and filter_shipregion:
                        dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                    else:
                        if len(cusVar) !=0 and filter_customer:
                            dfFilter=dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
                        else:
                            if len(plantVar)!=0:
                                dfFilter = dfFilter[dfFilter["Delivering_Plant"].isin(plantVar)]
                            else:
                                filter_plant = False
        else:
            dfFilter = data

        uniqueregion = numpy_object.unique(dfFilter['Region'])
        uniqueregion = numpy_object.vstack(uniqueregion)    
        uniqueplantdepot = numpy_object.unique(dfFilter['Delivering_Plant'])
        uniqueplantdepot = numpy_object.vstack(uniqueplantdepot)
        uniquepartnumber = numpy_object.unique(dfFilter['Material_Number'])
        uniquepartnumber = numpy_object.vstack(uniquepartnumber)
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + 10000,7)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + len(uniqueregion)-1,7)).Value = uniqueregion
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + 10000,6)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + len(uniquepartnumber)-1,6)).Value = uniquepartnumber
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,8),activeWorksheet_1.Cells(7 + 10000,8)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,8),activeWorksheet_1.Cells(7 + len(uniqueplantdepot)-1,8)).Value = uniqueplantdepot
        activeWorksheet_3.Range(activeWorksheet_3.Cells(4,23),activeWorksheet_3.Cells(4 + 60000,23)).ClearContents()
        
        if changeToOtherField == 0:
            InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Ship_Region',fieldValueList,1)
        else:
            InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Ship_Region',fieldValueList,1)

    # Unique Customer
    if field == 'Customer':
        if changeToOtherField == 0:
            dfFilter = data
            filter_customer = True
            if len(partVar)!=0 and filter_partno:
                dfFilter = data[data['Material_Number'].isin(partVar)]
                if len(regionVar)!=0 and filter_region:
                    dfFilter = dfFilter[dfFilter['Region'].isin(regionVar)]
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False            
                    else:
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False                    
                else:
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False            
                    else:
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False                                 
            else:
                if len(regionVar)!=0 and filter_region:
                    dfFilter = data[data['Region'].isin(regionVar)]
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False            
                    else:
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)] 
                            else:
                                filter_customer = False           
                else:
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = data[data['Delivering_Plant'].isin(plantVar)]
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False           
                    else:
                        if len(shipregionVar)!=0 and filter_shipregion:
                            dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            if len(cusVar)!=0:
                                dfFilter = data[data['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False    
                        else:
                            if len(cusVar)!=0:
                                dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            else:
                                filter_customer = False                   
                
        else:
            dfFilter = data            

        uniquecustomers = numpy_object.unique(dfFilter['Sold-To_Customerr_Name'])
        uniquecustomers = numpy_object.vstack(uniquecustomers)
        uniqueplant = numpy_object.unique(dfFilter['Delivering_Plant'])
        uniqueplant = numpy_object.vstack(uniqueplant)
        uniquepart = numpy_object.unique(dfFilter['Material_Number'])
        uniquepart = numpy_object.vstack(uniquepart)
        uniqueregion = numpy_object.unique(dfFilter['Region'])
        uniqueregion = numpy_object.vstack(uniqueregion)
        uniqueshipregion = numpy_object.unique(dfFilter['Ship-To_Region'])
        uniqueshipregion = numpy_object.vstack(uniqueshipregion)
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,10),activeWorksheet_1.Cells(7 + 10000,10)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,10),activeWorksheet_1.Cells(7 + len(uniquecustomers)-1,10)).Value = uniquecustomers
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + 10000,7)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + len(uniqueregion)-1,7)).Value = uniqueregion
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + 10000,6)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + len(uniquepart)-1,6)).Value = uniquepart
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,8),activeWorksheet_1.Cells(7 + 10000,8)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,8),activeWorksheet_1.Cells(7 + len(uniqueplant)-1,8)).Value = uniqueplant
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,9),activeWorksheet_1.Cells(7 + 10000,9)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,9),activeWorksheet_1.Cells(7 + len(uniqueshipregion)-1,9)).Value = uniqueshipregion
        activeWorksheet_3.Range(activeWorksheet_3.Cells(4,23),activeWorksheet_3.Cells(4 + 60000,23)).ClearContents()


    if field == 'Ship_Region':
        if changeToOtherField == 0:
            filter_shipregion = True
            dfFilter = data
            if len(partVar)!=0 and filter_partno:
                dfFilter = dfFilter[dfFilter['Material_Number'].isin(partVar)]
                if len(regionVar)!=0 and filter_region:
                    dfFilter = dfFilter[dfFilter['Region'].isin(regionVar)]
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False           
                    else:
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                else:
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False           
                    else:
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False
                                                  
            else:
                if len(regionVar)!=0 and filter_region:
                    dfFilter = dfFilter[dfFilter['Region'].isin(regionVar)]
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False           
                    else:
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 


                else:
                    if len(plantVar)!=0 and filter_plant:
                        dfFilter = dfFilter[dfFilter['Delivering_Plant'].isin(plantVar)]
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False           
                    else:
                        if len(cusVar)!=0 and filter_customer:
                            dfFilter = dfFilter[dfFilter['Sold-To_Customerr_Name'].isin(cusVar)]
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False 
                        else:
                            if len(shipregionVar)!=0:
                                dfFilter = dfFilter[dfFilter['Ship-To_Region'].isin(shipregionVar)]
                            else:
                                filter_shipregion = False        
        else:
            dfFilter = data


        uniqueshipregion = numpy_object.unique(dfFilter['Ship-To_Region'])
        uniqueshipregion = numpy_object.vstack(uniqueshipregion)
        uniqueplant = numpy_object.unique(dfFilter['Delivering_Plant'])
        uniqueplant = numpy_object.vstack(uniqueplant)
        uniquepart = numpy_object.unique(dfFilter['Material_Number'])
        uniquepart = numpy_object.vstack(uniquepart)
        uniqueregion = numpy_object.unique(dfFilter['Region'])
        uniqueregion = numpy_object.vstack(uniqueregion)
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + 10000,7)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,7),activeWorksheet_1.Cells(7 + len(uniqueregion)-1,7)).Value = uniqueregion
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + 10000,6)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,6),activeWorksheet_1.Cells(7 + len(uniquepart)-1,6)).Value = uniquepart
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,8),activeWorksheet_1.Cells(7 + 10000,8)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,8),activeWorksheet_1.Cells(7 + len(uniqueplant)-1,8)).Value = uniqueplant
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,9),activeWorksheet_1.Cells(7 + 10000,9)).ClearContents()
        activeWorksheet_1.Range(activeWorksheet_1.Cells(7,9),activeWorksheet_1.Cells(7 + len(uniqueshipregion)-1,9)).Value = uniqueshipregion
        activeWorksheet_3.Range(activeWorksheet_3.Cells(4,23),activeWorksheet_3.Cells(4 + 60000,23)).ClearContents()

        if changeToOtherField == 0:
            InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Customer',fieldValueList,1)
        else:
            InputChangeValidate(activeWorksheet_3,activeWorksheet_1,dfFilter,numpy_object,'Customer',fieldValueList,1)

def rgbToInt(rgb):
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return colorInt

def plant_wise_grouping(wholedata,data,Period,Dataframeobj):
   df_grouped = Dataframeobj.DataFrame()
   input_partlist = data['Material_Number'].unique()
   for i in input_partlist:
       df_material = data[data['Material_Number']==i]
       for y in df_material['Delivering_Plant'].unique():
           df_plant = df_material[df_material['Delivering_Plant']==y]
           if df_plant['Ship-To_Region'].nunique()>1:
                    df_copy = df_plant.copy()
                    df_copy = df_copy.drop_duplicates()
                    df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                    if Period == 'Monthly':
                            df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                            idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),wholedata['Actual_Goods_Movement_Date'].max(), freq='M')
                    elif Period == 'Quarterly':
                            df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                            idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Q')   
                    elif Period == 'Yearly':
                            df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                            idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Y')
                    elif Period =='Half Yearly':
                            df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                            idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='6M')
                    df_copy =df_copy.iloc[:,[2,5]]
                    df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                    df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                    df_copy.reset_index(drop=False,inplace=True)
                    df_copy['Material_Number'] = i
                    df_copy['Region'] = df_plant['Region'].unique()[0]
                    df_copy['Delivering_Plant'] = y
                    df_copy['Ship-To_Region'] = 'consolidate'
                    df_copy['Sold-To_Customerr_Name'] = 'consolidate'
                    df_copy.columns.values[0] = 'Period'
                    df_grouped = Dataframeobj.concat([df_grouped,df_copy],axis=0)
           else:
                for z in df_plant['Ship-To_Region'].unique():
                    df_ship_region = df_plant[df_plant['Ship-To_Region']==z]
                    if df_ship_region['Sold-To_Customerr_Name'].nunique()>1:
                        df_copy = df_ship_region.copy()
                        df_copy = df_copy.drop_duplicates()
                        df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                        if Period == 'Monthly':
                                df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                                idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),wholedata['Actual_Goods_Movement_Date'].max(), freq='M')
                        elif Period == 'Quarterly':
                                df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                                idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Q')   
                        elif Period == 'Yearly':
                                df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                                idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='Y')
                        elif Period =='Half Yearly':
                                df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                                idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),df_copy['Actual_Goods_Movement_Date'].max(), freq='6M')
                        df_copy =df_copy.iloc[:,[2,5]]
                        df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                        df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                        df_copy.reset_index(drop=False,inplace=True)
                        df_copy['Material_Number'] = i
                        df_copy['Region'] = df_plant['Region'].unique()[0]
                        df_copy['Delivering_Plant'] = y
                        df_copy['Ship-To_Region'] = 'consolidate'
                        df_copy['Sold-To_Customerr_Name'] = 'consolidate'
                        df_copy.columns.values[0] = 'Period'
                        df_grouped = Dataframeobj.concat([df_grouped,df_copy],axis=0)
                    else:
                        df_temp = period_wise_grouping(wholedata,df_ship_region,Period,Dataframeobj)
                        df_temp['Ship-To_Region'] = 'consolidate'
                        df_temp['Sold-To_Customerr_Name']='consolidate'
                        df_grouped = Dataframeobj.concat([df_grouped,df_temp],axis=0)
                
   return df_grouped

def period_wise_grouping(wholedata,data,Period,Dataframeobj):
   df_grouped = Dataframeobj.DataFrame()
   input_partlist = data['Material_Number'].unique()
   for i in input_partlist:
      df_material = data[data['Material_Number']==i]
      for y in df_material['Region'].unique():
         df_region = df_material[df_material['Region']==y]
         for z in df_region['Delivering_Plant'].unique():
            df_Plant = df_region[df_region['Delivering_Plant']==z]
            for n in df_Plant['Ship-To_Region'].unique():
                df_ship_region = df_Plant[df_Plant['Ship-To_Region']==n]
                for m in df_ship_region['Sold-To_Customerr_Name'].unique():
                    df_customer = df_ship_region[df_ship_region['Sold-To_Customerr_Name']==m]
                    df_copy = df_customer.copy()
                    df_copy = df_copy.drop_duplicates()
                    df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                    if Period == 'Monthly':
                        df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                        idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),wholedata['Actual_Goods_Movement_Date'].max(), freq='M')
                    elif Period == 'Quarterly':
                        df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                        idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today(), freq='Q')   
                    elif Period == 'Yearly':
                        df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                        idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today(), freq='Y')
                    elif Period =='Half Yearly':
                        df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                        idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today, freq='6M')

                    df_copy =df_copy.iloc[:,[2,5]]
                    df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                    df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                    df_copy.reset_index(drop=False,inplace=True)
                    df_copy['Material_Number'] = i
                    df_copy['Region'] = y
                    df_copy['Delivering_Plant'] = z
                    df_copy['Ship-To_Region'] = n
                    df_copy['Sold-To_Customerr_Name'] = m
                    df_copy.columns.values[0] = 'Period'
                    df_grouped = Dataframeobj.concat([df_grouped,df_copy],axis=0)
        
   return df_grouped

def global_grouping(wholedata,data,Period,Dataframeobj):
       df_grouped = Dataframeobj.DataFrame()
       input_partlist = data['Material_Number'].unique()
       for i in input_partlist:
            df_material = data[data['Material_Number']==i]
            for y in df_material['Region'].unique():
                df_region = df_material[df_material['Region']==y]
                df_copy = df_region.copy()
                df_copy = df_copy.drop_duplicates()
                df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                if Period == 'Monthly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),wholedata['Actual_Goods_Movement_Date'].max(), freq='M')
                elif Period == 'Quarterly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today(), freq='Q')   
                elif Period == 'Yearly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today(), freq='Y')
                elif Period =='Half Yearly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today, freq='6M')

                df_copy =df_copy.iloc[:,[2,5]]
                df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                df_copy.reset_index(drop=False,inplace=True)
                df_copy['Material_Number'] = i
                df_copy['Region'] = y
                df_copy['Delivering_Plant'] = 'consolidate'
                df_copy['Ship-To_Region'] = 'consolidate'
                df_copy['Sold-To_Customerr_Name'] = 'consolidate'
                df_copy.columns.values[0] = 'Period'
                df_grouped = Dataframeobj.concat([df_grouped,df_copy],axis=0)
       return df_grouped


def part_wise_grouping(wholedata,data,Period,Dataframeobj):
       df_grouped = Dataframeobj.DataFrame()
       input_partlist = data['Material_Number'].unique()
       for i in input_partlist:
                df_material = data[data['Material_Number']==i]
                df_copy = df_material.copy()
                df_copy = df_copy.drop_duplicates()
                df_copy = df_copy.sort_values(by=['Actual_Goods_Movement_Date'])
                if Period == 'Monthly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('M')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),wholedata['Actual_Goods_Movement_Date'].max(), freq='M')
                elif Period == 'Quarterly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Q')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today(), freq='Q')   
                elif Period == 'Yearly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('Y')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today(), freq='Y')
                elif Period =='Half Yearly':
                    df_copy['Actual_Goods_Movement_Date'] = Dataframeobj.to_datetime(df_copy['Actual_Goods_Movement_Date']).dt.to_period('6M')
                    idx = Dataframeobj.period_range(df_copy['Actual_Goods_Movement_Date'].min(),date.today, freq='6M')

                df_copy =df_copy.iloc[:,[2,5]]
                df_copy = df_copy.groupby(['Actual_Goods_Movement_Date']).Quantity_Delivered_Actual.sum().reset_index(drop=False)
                df_copy = df_copy.set_index('Actual_Goods_Movement_Date').reindex(idx, fill_value=0)
                df_copy.reset_index(drop=False,inplace=True)
                df_copy['Material_Number'] = i
                df_copy['Region'] = 'consolidate'
                df_copy['Delivering_Plant'] = 'consolidate'
                df_copy['Ship-To_Region'] = 'consolidate'
                df_copy['Sold-To_Customerr_Name'] = 'consolidate'
                df_copy.columns.values[0] = 'Period'
                df_grouped = Dataframeobj.concat([df_grouped,df_copy],axis=0)
       return df_grouped       

def create_forecast_table():
    conn = sqlite3.connect('test1.db')
    print("Opened database successfully")
    try:
        print('checkin')
        conn.execute('''CREATE TABLE if not exists Forecast
            (Id int NOT NULL PRIMARY KEY,
            Material_Number CHAR(10),
            Region CHAR(50),
            Delivering_Plant CHAR(50),
            Ship_To_Region CHAR(100),
            Sold_To_Customerr_Name CHAR(50),
            Lead_Time int,
            Standard_price_USD int,
            Service_Level int,
            Forecast_Buckets int,
            Forecast_Periods int,
            SafetyFactor real,
            Safety_stock int,
            ROP int,
            Max_Stock int,
            Churn_in_Dollar int,
            Delivery_Time char(20),
            Quantity_Delivered int,
            Predicted int
            );''')
        print('Creating Table!!!')
        conn.commit()
        print('Table created')
        conn.close() 
    except Exception as e:
        print(e)           
    #print('Successfully created the table!!!')

def delete_table():
    conn = sqlite3.connect('test1.db')

    conn.execute("drop table Override")
    conn.commit()
    print("Table Deleted")
    
def get_data():
    conn = sqlite3.connect('test1.db')
    print("Opened database successfully")
    cursor = conn.execute("select * from Override").fetchall()  
    return cursor 

def check_table_column():
    conn = sqlite3.connect('test1.db')
    cursor=conn.execute("PRAGMA table_info(Override)") 
    records = cursor.fetchall() 
    print(records)
    for row in records:
        print("Columns: ", row[1])

# check_table_column()        
# insert_data()

def insert_list(lst):
    conn = sqlite3.connect('test1.db')
    try:
        for i in lst:
            lis = [(i)]
            conn.executemany('INSERT INTO FORECAST VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);',lis) 
            print("New records updated------")
        conn.commit()
    except Exception as e:
        print(e) 
        return False    
    return True
 
# print(insert_list(),"------------")

def delete_all_records():
    conn=sqlite3.connect('test1.db')
    try:
        conn.execute('Delete From Override')
        conn.commit()
    except Exception as e:
        print(e)
        return False

def create_override_table():
    conn = sqlite3.connect('test1.db')
    print("Opened database successfully")
    conn.execute('''CREATE TABLE if not exists Override
            (Material_Number CHAR(50),
            Region CHAR(50),
            Delivering_Plant int,
            Ship_To_Region CHAR(100),
            Sold_To_Customerr_Name CHAR(50),
            Standard_price_USD int,
            Lead_Time int,
            Service_Level int,
            Forecast_Buckets int,
            Forecast_Periods int,
            SafetyFactor real,
            Safety_stock int,
            ROP int,
            Max_Stock int,
            Churn_in_Dollar int,
            Delivery_Time CHAR(100) NOT NULL,
            Quantity_Delivered int,
            Safety_stock_Override int,
            ROP_Override int,
            Max_Stock_Override int,
            Churn_in_Dollar_Override int,
            Date_Time CHAR(100),
            Predicted int
            );''')
    conn.commit()
    print("Override Table created ")    
    conn.close()  

def insert_to_override(lis):
    conn = sqlite3.connect('test1.db')
    try:
        for i in lis:
            lisVar = [(i)]
            conn.executemany('''INSERT INTO Override VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);''',lisVar) 
            print("New records updated------")
        conn.commit()
    except Exception as e:
        print(e)
        return False    
    return True

def check_table():
    conn = sqlite3.connect('test1.db')
    print("Opened database successfully")
    cursor = conn.execute("PRAGMA table_info(Override)")
    for i in cursor:
        print('Table existed !!!')
        return   
    create_override_table()