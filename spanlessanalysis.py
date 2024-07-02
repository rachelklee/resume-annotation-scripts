#spanless analysis
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
import pprint
from collections import OrderedDict
from datetime import date
import numpy as np

#import json
full_filepath = r"C:\Users\rachellee\AppData\Local\Microsoft\WindowsApps\spanlessoutput.xlsx"
df = pd.read_excel(full_filepath)

#slice by source
df_sliced_dict = {}

for source in df['source'].unique():
    df_sliced_dict[source] = df[df['source'] == source]

#dictionary to Dataframe
k1dict = {key: df_sliced_dict[key] for key in df_sliced_dict.keys() & {'Kaggle (1)'}}
k1val = list(k1dict.values())
k1= pd.DataFrame(k1val[0])

k2dict = {key: df_sliced_dict[key] for key in df_sliced_dict.keys() & {'Kaggle (2)'}}
k2val = list(k2dict.values())
k2= pd.DataFrame(k2val[0])

k3dict = {key: df_sliced_dict[key] for key in df_sliced_dict.keys() & {'Kaggle (3)'}}
k3val = list(k3dict.values())
k3= pd.DataFrame(k3val[0])

mdict = {key: df_sliced_dict[key] for key in df_sliced_dict.keys() & {'MITRE'}}
mval = list(mdict.values())
m= pd.DataFrame(mval[0])

#dataframe to excel

slicedannots=r"C:\Users\rachellee\Desktop\Resumes\Analysis\spanlessannotsbysource.xlsx"
with pd.ExcelWriter(slicedannots, engine='openpyxl', mode='a') as writer:
    k1.to_excel(writer, sheet_name='Kaggle (1)')
    k2.to_excel(writer, sheet_name='Kaggle (2)')
    k3.to_excel(writer, sheet_name='Kaggle (3)')
    m.to_excel(writer, sheet_name='MITRE')


k1count=k1['text'].value_counts()
k2count=k2['text'].value_counts()
k3count=k3['text'].value_counts()
mcount=m['text'].value_counts()

labelstats = r"C:\Users\rachellee\Desktop\Resumes\Analysis\spanlesslabelstats.xlsx"
with pd.ExcelWriter(labelstats, engine='openpyxl', mode='a') as writer:
    k1count.to_excel(writer, sheet_name='Kaggle (1)')
    k2count.to_excel(writer, sheet_name='Kaggle (2)')
    k3count.to_excel(writer, sheet_name='Kaggle (3)')
    mcount.to_excel(writer, sheet_name='MITRE')
