import tkinter as tk
from tkinter import filedialog
import pandas as pd

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue')

processState = 0

def getPriceBookExcel():
    global priceBookDataFrame
    global processState
    import_pb_file_path = filedialog.askopenfilename()
    priceBookDataFrame = pd.read_excel(import_pb_file_path)
   # print(priceBookDataFrame)
    pricebook_Excel["state"] = "disabled"
    processState = processState + 1
    if(processState == 2):
        process["state"]="active"



def getSurchargeExcel():
    global surchargeDataFrame
    global processState
    import_sc_file_path = filedialog.askopenfilename()
    surchargeDataFrame = pd.read_excel(import_sc_file_path)
   # print(surchargeDataFrame)
    surcharge_Excel["state"] = "disabled"
    processState = processState + 1
    if (processState == 2):
        process["state"] = "active"

def process():
    global processState
    global priceBookDataFrame

    pricebook_Excel["state"] = "active"
    surcharge_Excel["state"] = "active"
    process["state"]="disabled"
    processState=0
    print(surchargeDataFrame["Surcharge Percentage"])

    # iterate SurchrageDataFrame and prepare the map with key=Prefix.0Style+CoverSeries column and value is surcharge
    surchargeMap = {}
    nan = "NaN"
    for i in range(len(surchargeDataFrame)):
        if(str(surchargeDataFrame.loc[i, "Cover Series"] ==nan) and str(surchargeDataFrame.loc[i, "Cover Grade"] !=nan)):
            key = str(surchargeDataFrame.loc[i, "Item Number"])+ str(surchargeDataFrame.loc[i, "Cover Grade"])
        elif (str(surchargeDataFrame.loc[i, "Cover Series"] !=nan) and str(surchargeDataFrame.loc[i, "Cover Grade"] ==nan)):
            key = str(surchargeDataFrame.loc[i, "Item Number"])+ str(surchargeDataFrame.loc[i, "Cover Series"])
        elif(str(surchargeDataFrame.loc[i, "Cover Series"] ==nan) and str(surchargeDataFrame.loc[i, "Cover Grade"] ==nan)):
            key = str(surchargeDataFrame.loc[i, "Item Number"])
        key = str(surchargeDataFrame.loc[i, "Item Number"]) + str(surchargeDataFrame.loc[i, "Cover Series"])
        value=str(surchargeDataFrame.loc[i, "Surcharge Percentage"])
        surchargeMap[key.strip()]=value
    print(surchargeMap)

    #iterate PriceBookDataFrame and prepare and search key in SurchargeMap
    surchargeList=[]
    for i in range(len(priceBookDataFrame)):
        prefix = ""
        style = ""
        grade = ""
        pattern=""
        color = ""
        if(str(priceBookDataFrame.loc[i,"Prefix"]) != nan):
            prefix = str(priceBookDataFrame.loc[i,"Prefix"])
        if (str(priceBookDataFrame.loc[i, "Style"]) != nan):
            style = str(priceBookDataFrame.loc[i, "Style"])
        if (str(priceBookDataFrame.loc[i, "Grade"]) != nan):
            grade = str(priceBookDataFrame.loc[i, "Grade"])
        if (str(priceBookDataFrame.loc[i, "Pattern"]) != nan):
            pattern = str(priceBookDataFrame.loc[i, "Pattern"])
        if (str(priceBookDataFrame.loc[i, "Color"]) != nan):
            color = str(priceBookDataFrame.loc[i, "Color"])
        
        key=prefix+".0"+style+grade+pattern+color
        surcharge = surchargeMap.get(key.strip(),-1)
        if(surcharge == -1):
            key = prefix+".0"+style+grade+pattern
            surcharge = surchargeMap.get(key.strip(),-1)
            if(surcharge == -1):
                key = prefix+".0"+style+grade
                surcharge = surchargeMap.get(key.strip(),-1)
                if(surcharge == -1):
                    key = prefix+".0"+style
                    surcharge = surchargeMap.get(key.strip(), -1)
                    if(surcharge == -1):
                        surchargeList.append("")
                    else:
                        surchargeList.append(surcharge)
                else:
                    surchargeList.append(surcharge)
            else:
                surchargeList.append(surcharge)
        else:
            surchargeList.append(surcharge)

        print(str(key)+"="+str(surcharge))
    priceBookDataFrame=priceBookDataFrame.assign(Surcharge = surchargeList)
    print(priceBookDataFrame)
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    priceBookDataFrame.to_excel(export_file_path, index=False, header=True)





pricebook_Excel = tk.Button(text='Import Price Book Excel File', command=getPriceBookExcel, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
surcharge_Excel = tk.Button(text='Import Surcharge Book Excel File', command=getSurchargeExcel, bg='red', fg='white',
                               font=('helvetica', 12, 'bold'))
process = tk.Button(text='Process', command=process, bg='red', fg='white',font=('helvetica', 12, 'bold'))
process["state"] = "disabled"
canvas1.create_window(150, 150, window=pricebook_Excel)
canvas1.create_window(150, 200, window=surcharge_Excel)
canvas1.create_window(150, 250, window=process)
canvas1.pack()
root.mainloop()