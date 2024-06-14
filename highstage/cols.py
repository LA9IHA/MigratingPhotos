# Purpose: Prepare HighStage album with photos for  igration to Piwigo
# Pre requisites: Album.xlsx and Pic.xlsx is created from Highstage
# Licence: GNU 2.0
# Author: Ottar Kvindesland, 2024
# Reference: https://piwigo.miraheze.org/wiki/HighstageExport

class colsHighStage:

    def __init__(self):


        #self.homedir = "/Volumes/home/Oppgaver/transfer/"
        self.homedir = "/Volumes/T7Touch/album/transfer/"
        self.treedir = self.homedir + "tree4/"
        self.subdir = self.homedir + "source/"
        
        self.fInputAlbum = 'Album.xlsx'
        self.fOutputAlbum = 'Album1.xlsx' # When tested, this should be the same as Album.xlsx
        
        self.fInputPic = 'Pic.xlsx'
        self.fOutputPic = 'Pic1.xlsx'     # When tested, this should be the same as Pic.xlsx
        
        self.PiwigoPic = 'photos.xlsx'
        
        # Define top parent in HighStage. If blanks, update Album.xlsx and replace empty
        # parent album names with ZZZ and name topParent ZZZ.
        
        #self.topParent = 'SPC1064-1A'
        #self.topParent = 'ALBUM1732-1A'
        self.topParent = 'ALBUM1148-1A'
        
        self.path = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''] 
        self.path[0] = self.treedir
        self.MDerr = 'UFFDA'
        
        self.custom_image_extensions = ['.jpeg', '.png', '.jpg', '.gif']
        self.custom_substitutions = [('&', 'et'), ('(',''), (')',''), (' !', ''), ("'", ' '), (' ', '_'), (',', '_')]
        self.custom_video_mime_types = ['media', 'video']
        
        # Columns used in Album.xlsx
        self.ca = 0            # A = Id
        self.caItem = 1        # B = Item
        self.caDescription = 2 # C = Description
        self.caWorkspace = 3   # D = Workspace, i.e. access group
        self.caEventTime = 4   # E
        self.caEditBy = 5      # F
        self.caNote = 6        # G
        self.caInitdate = 7    # H
        self.caParentDoc = 8   # I
        self.caFileName = 9    # J
        self.caBareItem = 10   # K
        self.caSeq = 11        # L
        self.caMD5 = 12        # M
        self.caPiwigoId = 13   # N
        self.caLastAlbum = 19  # O
        
        self.colsAlbum = ['Line no', 'Item', 'Description', 'Workspace', 'EventTime', 'EditBy', 'Note', 'Initdate', 'ParentDoc', 'FileName', 'BareItem', 'Seq', 'MD5', 'PiwigoId', 'LastAlbum']
                
        # Columns used in Pic.xlsx
        self.cp = 0            # A = ID
        self.cpItem = 1        # B = Item
        self.cpDescription = 2 # C 
        self.cpWorkspace = 3   # D 
        self.cpEventTime = 4   # E 
        self.cpEditBy = 5      # F 
        self.cpNote = 6        # G 
        self.cpAlias = 7       # H 
        self.cpNote2 = 8       # I = History
        self.cpDate2 = 9       # J = First time storage
        self.cpDate3 = 10      # K
        self.cpExif = 11       # L
        self.cpInitdate = 12   # M = Date taken
        self.cpParentDoc = 13  # N
        self.cpFileName = 14   # O
        self.cpBareItem = 15   # P
        self.cpAlbumFile = 16  # Q
        self.cpSeq = 17        # R = Sequence
        self.cpFileError = 18  # S = Error message
        self.cpDest = 19       # T
        self.cpPiwigoId = 20   # U
        self.cpLastPic = 24    # V

        self.colsPic = ['Line no.', 'Item', 'Description', 'Workspace', 'EventTime', 'EditBy', 'Note', 'Alias', 'Note2', 'Date2', 'Date3', 'Exif', 'Initdate', 'ParentDoc', 'FileName', 'BareItem', 'AlbumFile', 'Seq', 'FileError', 'Dest', 'PiwigoId', 'LastPic']

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

