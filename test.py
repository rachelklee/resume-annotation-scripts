import json
from operator import index
from tabnanny import filename_only
from typing import final
import pandas as pd
import os
import warnings


def ingest_data(filename):
    full_filepath = r"C:\Users\rachellee\Desktop\Resumes\Reviewed"+"\\"+str(filename)
    print(full_filepath)
    data=json.load(open(full_filepath, encoding="utf8"))
    print(data.keys())
    df=pd.DataFrame(data["asets"])
    signal = data["signal"]
    #print(df.columns)
    #df=df.drop("hasSpan")
    df=df.drop(axis=0,index=[0,1])
    return filename,df,signal

def final_dataframe(filename,df,signal):
    finaldf=pd.DataFrame()
    for i in range(len(df["annots"][2:])):
        listofindexes=list(df["annots"][2:])[i]
        if len(listofindexes)==0:
            pass
        elif isinstance(listofindexes[0][0], str):
            pass
        else:
            for indexpair in listofindexes:
                text=signal[indexpair[0]:indexpair[1]]
                row = {'Filename':filename,'Label':list(df['type'][2:])[i],'text':text}
                warnings.filterwarnings("ignore")
                finaldf = finaldf.append(row, ignore_index=True)
    finaldf = finaldf.drop_duplicates()
    return finaldf

def to_excel_sheets():
    frames=[]
    for file in os.listdir(r"C:\Users\rachellee\Desktop\Resumes\Reviewed"):
        filename,df,signal=ingest_data(file)
        finaldf=final_dataframe(filename,df,signal)
        
        count=finaldf['Label'].value_counts()
        print(filename)
        print(count)

        outfile=r"C:\Users\rachellee\Desktop\Resumes\Analysis\testing.xlsx"

        with pd.ExcelWriter(outfile, engine='openpyxl', mode='a') as writer:
            count.to_excel(writer, sheet_name=filename)

    return count

sheets=to_excel_sheets()

