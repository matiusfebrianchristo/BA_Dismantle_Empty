import os, sys, time
import pandas as pd
from docxtpl import DocxTemplate
import warnings

#For move to currend path file
os.chdir(sys.path[0])

file_created = 0

def watermark():
    print("\n=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")
    print(f"       Success created {file_created} file  ")
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=")
    print("=-=     BUILD WITH \u2764\uFE0F BY MATIUS       =-=")
    print("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=\n")

#for ignore exception warnings
with warnings.catch_warnings():
    warnings.filterwarnings("ignore", category=UserWarning)
    
    #get data form excel
    try:
        sheet = pd.read_excel('Data_Excel.xlsx', 'Input')
    except Exception as message:
        print(message)
        
#Word Template
template = DocxTemplate('Template.docx')



#For create the document
def createDocument(data):
    try :
        template.render(data)
        
        #For fixed sitename form character "/" that make error
        sitename_fixed = ""
        temp = data['sitename'].split("/")
        if len(temp) == 1 :
            sitename_fixed = data['sitename']
        else :
            sitename_fixed = '_'.join(map(lambda a : a, temp))

        #Save the doucument
        template.save(f"result/doc/BA_Dismantle_{data['tower_owner']}_{data['site_id']}_{sitename_fixed}.docx")
        print(f"Document BA_Dismantle_{data['tower_owner']}_{data['site_id']}_{sitename_fixed} has been created!!")
        return "created"
        
    except Exception as message :
        print(message)



#For loop data in input excel
for data in sheet.to_numpy():
    contains = {
    'system_key' : data[0],
    'site_id' : data[1],
    'sitename' : data[2],
    'tower_owner' : data[3],
    'region' : data[4],
    'area' : data[5],
    'categoty': data[6],
    'pairing_site': data[7],
    'longtitude' : data[8],
    'latitude': data[9]
    }
    
    #For Create the documents
    if createDocument(contains) == 'created':
        file_created += 1

watermark()