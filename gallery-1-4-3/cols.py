# Purpose: Prepare Gallery album with photos for igestion to Piwigo
# Pre requisites: Album.xlsx and Pic.xlsx is created from Highstage
# Licence: GNU 2.0
# Author: Ottar Kvindesland, 2024
# Reference: https://piwigo.miraheze.org/wiki/GalleryExport

class cols:

    def __init__(self):
        
        self.testMode = True # Set to False if it is for a production run
        
        # Define top parent in Gallery 1.14.0. If blanks, update Album.xlsx and replace empty
        # parent album names with ZZZ and name topParent ZZZ.
        
        self.topParent = 'ottar_egne'

        #self.homedir = "/Volumes/home/Oppgaver/transfer/"
        self.homedir = "/Volumes/T7Touch/gallery/transfer/"
        self.treedir = self.homedir + self.topParent + "/"
        self.subdir = self.homedir + "source/"
        self.injectdir = self.homedir + "dest/"
        self.dbdumpdir = self.subdir + "dbdump/"
        self.errLog = self.subdir + 'x-port.log'
        
        self.fInputAlbum = 'Album.xlsx'
        self.fOutputAlbum = self.fInputAlbum
        
        self.fInputPic = 'Pic.xlsx'
        self.fOutputPic = self.fInputPic

        self.fInterAlbum = 'interAlbum.xlsx'

        if self.testMode:
            self.fOutputAlbum = 'Album1.xlsx' # When tested, this should be the same as Album.xlsx, set self.testMode to False
            self.fOutputPic = 'Pic1.xlsx'     # When tested, this should be the same as Pic.xlsx, set self.testMode to False
        
        self.fRefList = 'References.xlsx'
        
        self.PiwigoPic = 'images.xlsx'
        self.PiwigoAlbum = 'categories.xlsx'
        
        self.sqlFileName = 'metadata.sql'
        self.usersFileName = 'users.txt'
        self.userSqlFileName = 'users.sql'
        
        self.path = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''] 
        self.path[0] = self.treedir
        self.MDerr = 'UFFDA'
        
        self.custom_image_extensions = ['.jpeg', '.png', '.jpg', '.gif']
        self.custom_substitutions = [('&', 'et'), ('(',''), (')',''), (' !', ''), ("'", ' '), (' ', '_'), (',', '_')]
        self.custom_video_mime_types = ['media', 'video']
        
        # Columns used in Album.xlsx
        self.caJson = 0         # A
        self.caAlbumName = 1    # B
        self.caItem = 2         # C
        self.caEditBy = 3       # D
        self.caParentDoc = 4    # E
        self.caDescription = 5  # F
        self.caEventTime = 6    # G
        self.caInitdate = 7     # H
        self.ca = 8             # I
        self.caFileName = 9     # J
        self.caBareItem = 10    # K
        self.caSeq = 11         # L
        self.caMD5 = 12         # M
        self.caPiwigoId = 13    # N
        self.caAlbumImg = 14    # O
        self.caAlbumPath = 15   # P
        self.caParentPwgId = 16 # S
        self.caLastAlbum = 20   # T
        
        self.colsAlbum = ['Json', 'AlbumName', 'Item', 'EditBy', 'ParentDoc', 'Description', 'EventTime', 'Initdate', 'ca', 'FileName', 'BareItem', 'Seq', 'MD5', 'PiwigoId', 'AlbumImg', 'AlbumPath', 'ParentPwgId', '', '', 'LastAlbum']
        
        # Columns used in Pic.xlsx
        self.cpDescription = 0  # A 
        self.cpFileType = 1     # B
        self.cpItem = 2         # C  Item
        self.cpFileName = 3     # D
        self.cpParentDoc = 4    # E
        self.cpEditBy = 5       # F 
        self.cpComment = 6      # G
        self.cpAlbumFile = 7    # H  Also albumFile for albums when file_name (3-D) is empty
        self.cpInitdate = 8     # I  Date taken
        self.cpKeyWord = 9      # J
        self.cpSeq = 10         # K  Also for albums
        self.cp = 11            # L  ID
        self.cpBareItem = 12    # M
        self.cpFileError = 13   # N  Error message
        self.cpDest = 14        # O
        self.cpPiwigoId = 15    # P
        self.cpMigrInfo = 16    # Q
        self.cpPath = 17        # R
        self.cpParentPwgId = 18 # S
        self.cpAlbImgIdId = 19  # T
        self.cpLastPic = 24     # V

        self.colsPic = ['Description', 'FileType', 'Item', 'FileName', 'ParentDoc', 'EditBy', 'Comment', 'AlbumFile', 'Initdate', 'KeyWord', 'Seq', 'cp', 'BareItem', 'FileError', 'Dest', 'PiwigoId', 'MigrInfo', 'Path', 'ParentPwgId', 'AlbImgIdId', '', '', '', 'LastPic']

        # Columns used in References.xlsx
        self.rpParent = 4 # E
        self.rpChild = 6 # G
        self.rpSeq = 9 # J
        self.rpLastCol = 10 # K
        
        # Columns used in photos.xlsx
        self.ppid =  0                   # A
        self.ppfile = 1                  # B
        self.ppdate_available = 2        # C 
        self.ppdate_creation = 3         # D
        self.ppname = 4                  # E
        self.ppcomment =  5              # F
        self.ppauthor = 6                # G
        self.pphit = 7                   # H
        self.ppfilesize = 8              # I
        self.ppwidth = 9                 # J
        self.ppheight = 10               # K
        self.ppcoi = 11                  # L
        self.pprepresentative_ext = 12   # M
        self.ppdate_metadata_update = 13 # N
        self.pprating_score = 14         # O
        self.pppath = 15                 # P
        self.ppstorage_category_id = 16  # Q
        self.pplevel = 17                # R
        self.ppmd5sum = 18               # W
        self.ppadded_by = 19             # T
        self.pprotation = 20             # U
        self.pplatitude = 21             # V
        self.pplongitude = 22            # W
        self.pplastmodified = 23         # X
        self.pplastcol = 24              # Y
        
        # Coloumns from categories.xlsx
        self.paId = 0                        # A
        self.paname = 1                      # B
        self.paid_uppercat = 2               # C 
        self.pacomment = 3                   # D
        self.padir = 4                       # E
        self.parank = 5                      # F
        self.paExample = 6                   # G
        self.parepresentative_picture_id = 7 # H
        self.pauppercats = 8                 # I
        self.pacommentable = 9               # J
        self.paglobal_rank = 10              # K
        self.paimage_order = 11              # L
        self.papermalink = 12                # M
        self.palastmodified = 13             # N
        self.palastcol = 15                  # P
        
        # Columns from interAlbum.xlsx
        self.iaDescription = 0 # A
        self.iaFileType = 1    # B
        self.iaItem = 2        # C
        self.iaFileName = 3    # D
        self.iaParentDoc = 4   # E
        self.iaLastCol = 5     # F