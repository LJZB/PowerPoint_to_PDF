# -*- coding: utf-8 -*-
"""
Created on Thu May 26 13:37:04 2022

@author: lzuluaga
"""
import aspose.slides as ass
import os

path = os.getcwd()
i=0
try:
    os.system('cls')
    for file_name in os.listdir(path): # Find all the files inside the folder
        # PPTtoPDF(file_name, outputFileName)
        if file_name[-4:] == "pptx":
            size = len(file_name)
            # Load presentation
            pres = ass.Presentation(file_name)
            file_name = file_name[:size-5]
            # Convert PPTX to PDF
            file_name = file_name + ".pdf"
            pres.save(file_name, ass.export.SaveFormat.PDF)
            i+=1
    print('Number of files converted:',i)
        
except:
    print('ERROR')