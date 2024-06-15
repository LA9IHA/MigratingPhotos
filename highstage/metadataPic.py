import datetime
from openpyxl import Workbook
import openpyxl
import os
import shutil
import hashlib

from openpyxl.descriptors.base import DateTime
from openpyxl.reader.excel import load_workbook
from os import listdir

# See cols.py header for info.
# (C) 2024: Ottar Kvindesland, Licence: GPL 2.0
# Purpose: Build reference sequence numbers in album and files

from cols import colsHighStage

class addSeqs(colsHighStage):

    def __init__(self):
        
        super().__init__()
        
        self.creationDate = datetime.datetime(1980, 1, 1, 1, 0)
        
        picfile = self.subdir + self.fInputPic
        self.pic_wb = load_workbook(filename=picfile)
        self.wp = self.pic_wb.worksheets[0]
        
        albumfile = self.subdir + self.fInputAlbum
        self.album_wb = load_workbook(filename=albumfile)
        self.wa = self.album_wb.worksheets[0]
                
        self.refPics()
        self.refAlbums()
        
    def refPics(self):
        n = 0
        c = self.cpParentDoc+1
        bi = self.cpBareItem+1
        fn = self.cpFileName+1
        err = self.cpFileError+1
        sec = 0
        for pic in self.wp.iter_rows(min_row=1, max_row=self.wp.max_row, min_col=1, max_col=self.cpLastPic, values_only=False):
            n+=1
            parent = self.wp.cell(row=n, column=c).value
            bareItem = self.wp.cell(row=n, column=bi).value
            
            if parent is not None and bareItem is not None:
                seq = self.getSeq(parent, bareItem)
                if (seq > 0):
                    self.wp.cell(row=n, column=self.cpSeq+1).value = seq
                
            if (n % 100 == 0):
                print (n, ' pictures done')
        
        i = 1
        for col in self.colsPic:
            self.wp.cell(row=1, column=i).value = col
            i = i + 1
            
        self.pic_wb.save(self.subdir + self.fOutputPic)
        
    def refAlbums(self):
        n = 0
        c = self.caParentDoc+1
        bi = self.caBareItem+1
        sec = 0
        
        for album in self.wa.iter_rows(min_row=1, max_row=self.wa.max_row, min_col=1, max_col=self.caLastAlbum, values_only=False):
            n+=1
            parent = self.wa.cell(row=n, column=c).value
            bareItem = self.wa.cell(row=n, column=bi).value
            
            if parent is not None and bareItem is not None:
                seq = self.getSeq(parent, bareItem)
                if (seq > 0):
                    self.wa.cell(row=n, column=self.caSeq+1).value = seq

            if (n % 100 == 0):
                print (n, ' albums done')
                
        i = 1
        for col in self.colsAlbum:
            self.wa.cell(row=1, column=i).value = col
            i = i + 1
            
        self.album_wb.save(self.subdir + 'Album1.xlsx')
    
a = addSeqs()
