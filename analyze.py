import pandas as pd
from openpyxl import load_workbook
import numpy as np
import xlsxwriter
import json
from operator import index
from tabnanny import filename_only
from typing import final
import os
import warnings


#import excel
def process_excel(filename):
    full_filepath = r"C:\Users\rachellee\Desktop\Resumes\Analysis"+"\\"+str(filename)
    df = pd.read_excel(full_filepath)
    return df

#count frequencies
def annotations_per_resume(df):
    count=df['Filename'].value_counts()
    #print(count)
    return count

def frequency_of_annotations(df):
    count2=df['Label'].value_counts()
    #print(count2)
    return count2

#process json files
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

#json to dataframe
#label = list(df['type'][0:])[i]
def final_dataframe(filename,df,signal):
    finaldf=pd.DataFrame()
    for i in range(len(df["annots"][0:])):
        listofindexes=list(df["annots"][0:])[i]
        if len(listofindexes)==0:
            continue
        elif list(df['type'][0:])[i] == "SEGMENT":
            continue
        elif isinstance(listofindexes[0][0], str):
            continue
        else:
            for indexpair in listofindexes:
                text=signal[indexpair[0]:indexpair[1]]
                warnings.filterwarnings("ignore")
                row = {'Filename':filename,'Label':list(df['type'][0:])[i],'text':text}
                finaldf = finaldf.append(row, ignore_index=True)
    #print(finaldf)
    finaldf = finaldf.drop_duplicates()
    return finaldf

#dataframe to excel
def to_excel_sheets():
    frames=[]
    for file in os.listdir(r"C:\Users\rachellee\Desktop\Resumes\Reviewed"):
        filename,df,signal=ingest_data(file)
        finaldf=final_dataframe(filename,df,signal)
        #print(finaldf)
        frames.append(finaldf)
        #print(frames)
        count=finaldf['Label'].value_counts()
        #print(filename)
        #print(count)

        outfile=r"C:\Users\rachellee\Desktop\Resumes\Analysis\spannedbyfile.xlsx"

        with pd.ExcelWriter(outfile, engine='openpyxl', mode='a') as writer:
            count.to_excel(writer, sheet_name=filename)

    return count

#run function for individual counts
sheets=to_excel_sheets()

#run function total counts
infile = r"C:\Users\rachellee\Desktop\Resumes\Analysis\spanned_annotations.xlsx"
df=process_excel("spanned_annotations.xlsx")
annotationsperresume=annotations_per_resume(df)
#print(annotationsperresume)
frequencyofannotations=frequency_of_annotations(df)
#print(frequencyofannotations)

#output total counts
outfile=r"C:\Users\rachellee\Desktop\Resumes\Analysis\spannedstats.xlsx"

with pd.ExcelWriter(outfile, engine='openpyxl', mode='a') as writer:
    annotationsperresume.to_excel(writer, sheet_name='Annotations per Resume')
    frequencyofannotations.to_excel(writer, sheet_name='Frequency of Annotations')

#annotationsperresume.to_excel(outfile, sheet_name='Annotations per Resume')
#frequencyofannotations.to_excel(outfile, sheet_name='Frequency of Annotations')