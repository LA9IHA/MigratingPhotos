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

# Purpose: Prepare HighStage album with photos for  igration to Piwigo
# Pre requisites: Album.xlsx and Pic.xlsx is created from Highstage
# Licence: GNU 2.0
# Author: Ottar Kvindesland, 2024
# Reference: https://piwigo.miraheze.org/wiki/HighstageExport

class getId(colsHighStage):

    def __init__(self):
        super().__init__()
       
        picfile = self.subdir + self.fInputPic
        self.pic_wb = load_workbook(filename=picfile)
        self.wp = self.pic_wb.worksheets[0]
        
        albumfile = self.subdir + self.fInputAlbum
        self.album_wb = load_workbook(filename=albumfile)
        self.wa = self.album_wb.worksheets[0]
        
        self.getPicsId()
        self.pic_wb.save(self.subdir + self.fOutputPic)
        
    def getPicsId(self):
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
                
    def testFile(self, p):
        dp = 0
        smallFiles = 0
        goodSuffixes = 0
        
        for fil in os.listdir(p):
            ppath = os.path.join(p, fil)
            if os.path.isfile(ppath):
                fsize = os.path.getsize(ppath)
                
                if ppath.endswith('doc_pic.jpg'):
                    dp = dp + 1
                    if fsize < 750:
                        smallFiles = smallFiles + 1
                else:
                    if fsize < 1500:
                        smallFiles = smallFiles + 1
                    if (any(ppath.lower().endswith(filetype) for filetype in self.custom_image_extensions) == True):
                        goodSuffixes = goodSuffixes + 1
            
        #print('GS: ', goodSuffixes)
        return (( dp == 1) and (smallFiles == 0)  and (goodSuffixes == 1) )
    
a = getId()
