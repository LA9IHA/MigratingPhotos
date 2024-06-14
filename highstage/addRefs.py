import datetime
from openpyxl import Workbook
import openpyxl
import os
import shutil
import hashlib

from openpyxl.descriptors.base import DateTime
from openpyxl.reader.excel import load_workbook
from os import listdir

from colsHighStage import colsHighStage

# Purpose: Prepare HighStage album with photos for Migration to Piwigo
# Pre requisites: Album.xlsx and Pic.xlsx is created from Highstage
# Licence: GNU 2.0
# Author: Ottar Kvindesland, 2024
# Reference: https://piwigo.miraheze.org/wiki/HighstageExport

class addSeqs(colsHighStage):

    def __init__(self, h, t, s):
        
        super().__init__()
        
        self.creationDate = datetime.datetime(1980, 1, 1, 1, 0)
        
        picfile = self.subdir + self.fInputPic
        self.pic_wb = load_workbook(filename=picfile)
        self.wp = self.pic_wb.worksheets[0]
        
        albumfile = self.subdir + self.fInputAlbum
        self.album_wb = load_workbook(filename=albumfile)
        self.wa = self.album_wb.worksheets[0]
        
        referencesfile = self.subdir + 'References.xlsx'
        self.references_wb = load_workbook(filename=referencesfile)
        self.wr = self.references_wb.worksheets[0]
                
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

            seq = self.getSeq(parent, bareItem)
            if (seq > 0):
                self.wp.cell(row=n, column=self.cpSeq+1).value = seq
                
            if (n % 100 == 0):
                print (n, ' pictures done')
        
        i = 1
        for col in self.colsPic:
            self.wp.cell(row=1, column=i).value = col
            i = i + 1
            
        self.pic_wb.save(self.subdir + 'Pic1.xlsx')
        
    def refAlbums(self):
        n = 0
        c = self.caParentDoc+1
        bi = self.caBareItem+1
        sec = 0
        
        for album in self.wa.iter_rows(min_row=1, max_row=self.wa.max_row, min_col=1, max_col=self.caLastAlbum, values_only=False):
            n+=1
            parent = self.wa.cell(row=n, column=c).value
            bareItem = self.wa.cell(row=n, column=bi).value
            
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

    def getSeq(self, p, c):
        
        p = p.split('-')[0]
        c = c.split('-')[0]
        
        n = 0
        cp = self.rpParent+1
        cc = self.rpChild+1
        cs = self.rpSeq+1
        hits = 0
        s = 0
        
        for ref in self.wr.iter_rows(min_row=1, max_row=self.wr.max_row, min_col=1, max_col=self.rpLastCol, values_only=False):
            n+=1
            parent = self.wr.cell(row=n, column=cp).value.split('-')[0]
            child = self.wr.cell(row=n, column=cc).value.split('-')[0]
            
            if (p == parent and c == child):
                hits = hits + 1
                s = int(self.wr.cell(row=n, column=cs).value)
        return s


homedir = "/Volumes/home/Oppgaver/hstransfer/" 
treedir = homedir + "tree/"
subdir = homedir + "source/"
    
a = addSeqs(homedir, treedir, subdir)
