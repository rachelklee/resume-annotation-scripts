#variance
#goal: to get variance per lable across N resumes
from fileinput import filename
import pandas as pd
import math
import xlsxwriter

N=118

#import excel
stats = r"C:\Users\rachellee\Desktop\Resumes\Analysis\spannedstats.xlsx"
statsdf = pd.read_excel(stats, sheet_name='Frequency of Annotations')

spannedannots = r"C:\Users\rachellee\Desktop\Resumes\Analysis\spanned_annotations.xlsx"
datadf = pd.read_excel(spannedannots)


labellist = list(statsdf[statsdf.columns[0]])[:-1]
averages = list(statsdf['Average'])[:-1]
#print(labellist)
#print(averages)

#create dictionary
labelaverages = dict(zip(labellist,averages))
allvariance = dict()
standarddeviation = dict()
#print(labelaverages)

allresumes = list(datadf['Filename'].unique())
#print(allresumes)

for label,mu in labelaverages.items():
    #print(label,mu)
    #print(label,mu)
    intermediate_value = 0
    for resumename in allresumes:
        slicedf = datadf.loc[datadf['Filename'] == resumename]
        x=len(slicedf.loc[slicedf['Label']==label])
        intermediate_value += (x-mu)**2

    label_variance = intermediate_value/N
    allvariance[label] = label_variance
    standard_deviation = math.sqrt(intermediate_value/N)
    standarddeviation[label] = standard_deviation

#print("variance:", allvariance)
#print("standarddeviation", standarddeviation)
#print("labelaverages", labelaverages)

variance = pd.DataFrame(list(allvariance.items()))
sdev = pd.DataFrame(list(standarddeviation.items()))

#output total variance
outfile=r"C:\Users\rachellee\Desktop\Resumes\Analysis\spannedstats.xlsx"

with pd.ExcelWriter(outfile, engine='openpyxl', mode='a') as writer:
    variance.to_excel(writer, sheet_name='Variance')
    sdev.to_excel(writer, sheet_name='Standard Deviation')