import openpyxl
from openpyxl import Workbook
#from openpyxl.descriptors.base import DateTime
from openpyxl.reader.excel import load_workbook
import os
import shutil
import datetime as dt
from datetime import datetime, timedelta

# See cols.py header for info.
# (C) 2024: Ottar Kvindesland, Licence: GPL 2.0
# Purpose: Initiate metadata files Album.xlsx and Pic.xlsx


from cols import cols

class init(cols):

    def __init__(self):
        
        super().__init__()

        self.path_depth = 0
        
        if os.path.exists(self.treedir):
            shutil.rmtree(self.treedir)
        os.makedirs(self.treedir)
        
        for filname in os.listdir(self.treedir):
            filepath = os.path.join(self.treedir, filname)
            try:
                if os.path.isfile(filepath) or os.path.islink(filepath):
                    os.unlink(filepath)
                elif os.path.isdir(filepath):
                    shutil.rmtree(filepath)
            except Exception as e:
                print(f'Could not purge {filepath}')
                
        picfile = self.subdir + self.fInputPic
        self.pic_wb = load_workbook(filename=picfile)
        self.wp = self.pic_wb.worksheets[0]
            
        albumfile = self.subdir + self.fInputAlbum
        self.album_wb = load_workbook(filename=albumfile)
        self.wa = self.album_wb.worksheets[0]
        
        if self.firstTime:
            self.addSeqPic()
            self.addSeqAlbum()
        else:
            self.lines = 1
            #self.albums2path('0', self.treedir)
            self.pic2path()
            
        self.album_wb.save(self.subdir + self.fOutputAlbum)
        self.pic_wb.save(self.subdir + self.fOutputPic)
        
    def addSeqPic (self):
        n = 1
        for row in self.wp.iter_rows(min_row=1, max_row=self.wp.max_row, min_col=1, max_col=self.cpLastPic, values_only=True):
            if n > 1:
                self.wp.cell(row=n, column=self.cp+1).value = n-1
            if (n % 100) == 0:
                print('Indexed ', n, ' photos')
            n += 1
        c = 1
        for col in self.colsPic:
            self.wp.cell(row=1, column=c).value = col
            c+=1

        print ('  Done, saving at ', self.subdir + self.fOutputPic)
        
    def addSeqAlbum (self):
        
        n = 1
        for album in self.wa.iter_rows(min_row=1, max_row=self.wa.max_row, min_col=1, max_col=self.caLastAlbum, values_only=True):
            if n > 1:
                self.wa.cell(row=n, column=self.ca+1).value = n-1
                start_date = dt.datetime.strptime('1970-01-01 00:00:00', '%Y-%m-%d %H:%M:%S')
                event_time = self.wa.cell(row=n, column=self.caEventTime+1).value
                seconds_since_start = (event_time - start_date).total_seconds()
                self.wa.cell(row=n, column=self.caEventTime+1).value = start_date + dt.timedelta(seconds=seconds_since_start)
            if (n % 100) == 0:
                print('Indexed ', n, ' albums')
            n += 1
        a = 1
        for col in self.colsAlbum:
            self.wa.cell(row=1, column=a).value = col
            a = a + 1
            
        print ('  Done, saving at ', self.subdir + self.fOutputAlbum)

    def albums2path (self, parent, path):
        
        for album in self.wa.iter_rows(min_row=1, max_row=self.wa.max_row, min_col=1, max_col=self.caLastAlbum, values_only=True):
            if str(album[self.caParentDoc]) == parent:
                if (self.lines % 100) == 0:
                    print ('Added path to ', self.lines, ' albums')
                self.lines += 1
                self.wa.cell(row=album[self.ca]+1, column=self.caAlbumPath+1).value = path + album[self.caItem] + '/'
                self.albums2path(album[self.caItem], path + album[self.caItem] + '/')
    
    def pic2path (self):
        
        self.lines = 1
        for pic in self.wp.iter_rows(min_row=1, max_row=self.wp.max_row, min_col=1, max_col=self.cpLastPic, values_only=True):
            if self.lines < 50:
                for album in self.wa.iter_rows(min_row=1, max_row=self.wa.max_row, min_col=1, max_col=self.caLastAlbum, values_only=True):
                    if str(pic[self.cpParentDoc]) == str(album[self.caItem]):
                        #print (album[self.caAlbumPath])
                        print (album[self.caAlbumPath] + pic[self.cpItem] + '.' + pic[self.cpFileType])
                        #self.wp.cell(row=int(pic[self.cp]+1), column=self.cpPath+1).value = album[self.caAlbumPath] + pic[self.cpFileType] + '.' + pic[self.cpItem]
                self.lines += 1
                if (self.lines % 100) == 0:
                    print ('Added path to ', self.lines, ' photos')
           
    
a = init()