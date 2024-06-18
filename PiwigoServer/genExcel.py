import datetime
#import openpyxl
#from openpyxl import Workbook
import os
import shutil
import hashlib
import pandas as pd
import fnmatch

import mysql.connector
#pip3 install mysql-connector-python

from openpyxl.descriptors.base import DateTime
from openpyxl.reader.excel import load_workbook
from os import listdir

# Purpose: Check the integrity of the media files
# Licence: GNU 2.0
# Copyright: Ottar Kvindesland, 2024

class genExcels ():

    def __init__(self):
        super().__init__()
        
        self.dbName = 'album'
        self.host='localhost'
        self.user='ottar'
        self.password='ottar'
        self.csvDir='/home/ottar/'
            
        tables = self.getCSVfiles()
        for table in tables:
            self.genXlsx (table)
                
    def getCSVfiles(self):
        
        results = [file for file in os.listdir(self.csvDir) if fnmatch.fnmatch(file, '*.csv')]
        r = []
        for row in results:
            r.append(row)
            
        return r

    def genXlsx (self, tableName):
        
        table = tableName[0]      
        if os.path.exists(file_path):
	        df = pd.read_csv(self.csvDir + tableName)
	        df.to_excel(self.csvDir + tableName, index=False)
        	os.remove(self.csvDir + tableName)

a = genExcels()


