import datetime
#import openpyxl
#from openpyxl import Workbook
import os
import shutil
import hashlib

import mysql.connector
#pip3 install mysql-connector-python

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
            
        tables = self.getAllTableNames()
        for table in tables:
            self.dumpTable (table)
        
        #picfile = self.subdir + self.fInputPic
        #self.pic_wb = load_workbook(filename=picfile)
        #self.wp = self.pic_wb.worksheets[0]
        
        #self.checkPics()
        #self.pic_wb.save(self.subdir + self.fOutputPic)
                
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
        c2 = "SET @sql = CONCAT('SELECT ', @col_names, ' UNION ALL SELECT * FROM " + table + " INTO OUTFILE \\'" + table + ".csv\\' FIELDS TERMINATED BY \\\',\\\' ENCLOSED BY \\\'\"\\\' LINES TERMINATED BY \\\'\n\\\'');"
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
        
    def checkPics(self):
        n = 0
        c = self.cpParentDoc+1
        bi = self.cpBareItem+1
        fn = self.cpFileName+1
        err = self.cpFileError+1
        for pic in self.wp.iter_rows(min_row=1, max_row=self.wp.max_row, min_col=1, max_col=self.cpLastPic, values_only=False):
            n+=1
            parent = self.wp.cell(row=n, column=c).value
            bareItem = self.wp.cell(row=n, column=bi).value
            fileName = self.wp.cell(row=n, column=fn).value
            fpath = self.subdir + 'PHOTOS/' + bareItem + '/' + bareItem + '/'
            fthumb = fpath + 'doc_pic.jpg'
            ffull = fpath + fileName
            
            if (bareItem.lower() != 'name'):
                if not(self.testFile(fpath)):
                    self.wp.cell(row=n, column=self.cpFileError).value = 'File Error'
                    print('ERROR: ' + fpath)
                
        

  
a = extractTables()


