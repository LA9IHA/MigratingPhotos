import datetime
#import openpyxl
#from openpyxl import Workbook
import os
import shutil
import hashlib
import pandas as pd
import fnmatch

import mysql.connector

from openpyxl.descriptors.base import DateTime
from openpyxl.reader.excel import load_workbook
from os import listdir

# Purpose: Check the integrity of the media files
# Licence: GNU 2.0
# Copyright: Ottar Kvindesland, 2024

class extractTables ():

    def __init__(self):
        super().__init__()
        
        self.dbName = 'album'
        self.host='localhost'
        self.user='ottar'
        self.password='ottar'
        self.migrateDir = '/home/ottar/migration/tables/'
            
        #tables = self.getAllTableNames()
        #for table in tables:
        #    self.dumpTable (table)
        
        tables = self.getCSVfiles()
        for table in tables:
            self.genXlsx (table)
        
                
    def getAllTableNames(self):
        
        conn = mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.dbName
            )
        cursor = conn.cursor()
        cursor.execute('show tables')
        results = cursor.fetchall()
        r = []
        for row in results:
            r.append(row)
            
        return r

    def dumpTable (self, tableName):
        
        table = tableName[0]
        
        print ('Printer: ', table)

        c1 = "SET @col_names = (SELECT GROUP_CONCAT(CONCAT('\\'', COLUMN_NAME, '\\'')) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" + table + "' AND TABLE_SCHEMA = '" + self.dbName + "');"
        c2 = "SET @sql = CONCAT('SELECT ', @col_names, ' UNION ALL SELECT * FROM " + table + " INTO OUTFILE \\'" +  self.migrateDir + table + ".csv\\' FIELDS TERMINATED BY \\\',\\\' ENCLOSED BY \\\'\"\\\' LINES TERMINATED BY \\\'\n\\\'');"
        c3 = 'PREPARE stmt FROM @sql;'
        c4 = 'EXECUTE stmt;'
        c5 = 'DEALLOCATE PREPARE stmt;'

        conn = mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.dbName
            )
        cursor = conn.cursor()
        cursor.execute(c1)
        cursor.execute(c2)
        cursor.execute(c3)
        cursor.execute(c4)
        cursor.execute(c5)
        
    def getCSVfiles(self):
        
        results = [file for file in os.listdir(self.migrateDir) if fnmatch.fnmatch(file, '*.csv')]
        r = []
        for row in results:
            r.append(row)
            
        return r

    def genXlsx (self, tableName):
        
        table = tableName[0]
        csvFile = ''
        if os.path.exists(file_path):
            csvFile = self.migrateDir + tableName + '.csv'
            df = pd.read_csv(csvFile)
            df.to_excel(self.migrateDir + tableName + '.xlsx', index=False)
            os.remove(csvFile)
  
a = extractTables()


