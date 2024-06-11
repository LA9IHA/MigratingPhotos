# Purpose: Prepare HighStage album with photos for  igration to Piwigo
# Pre requisites: Album.xlsx and Pic.xlsx is created from Highstage
# Licence: GNU 2.0
# Author: Ottar Kvindesland, 2024
# Reference: https://piwigo.miraheze.org/wiki/HighstageExport

class colsHighStage:

    def __init__(self):

        # Columns used in Album.xlsx
        self.ca = 0 # A = Id
        self.caItem = 1 # B = Item
        self.caDescription = 2 # C = Description
        self.caWorkspace = 3 # D = Workspace, i.e. access group
        self.caEventTime = 4 # etc....
        self.caEditBy = 5
        self.caNote = 6
        self.caInitdate = 7
        self.caParentDoc = 8
        self.caFileName = 9
        self.caBareItem = 10
        self.caSeq = 11
        self.caMD5 = 12
        self.caPiwigoId = 13
        self.caLastAlbum = 14
        
        self.colsAlbum = ['Line no', 'Item', 'Description', 'Workspace', 'EventTime', 'EditBy', 'Note', 'Initdate', 'ParentDoc', 'FileName', 'BareItem', 'Seq', 'MD5', 'PiwigoId', 'LastAlbum']
                
        # Columns used in Pic.xlsx
        self.cp = 0 # A = ID
        self.cpItem = 1 # B = Item
        self.cpDescription = 2 # etc ...
        self.cpWorkspace = 3
        self.cpEventTime = 4
        self.cpEditBy = 5
        self.cpNote = 6
        self.cpAlias = 7
        self.cpNote2 = 8 # History
        self.cpDate2 = 9  # First time storage
        self.cpDate3 = 10
        self.cpExif = 11
        self.cpInitdate = 12 # Date taken
        self.cpParentDoc = 13
        self.cpFileName = 14
        self.cpBareItem = 15
        self.cpAlbumFile = 16
        self.cpSeq = 17  # Q Sequence
        self.cpFileError = 18 # R - Error message
        self.cpDest = 19
        self.cpPiwigoId = 20
        self.cpLastPic = 21

        self.colsPic = ['Line no.', 'Item', 'Description', 'Workspace', 'EventTime', 'EditBy', 'Note', 'Alias', 'Note2', 'Date2', 'Date3', 'Exif', 'Initdate', 'ParentDoc', 'FileName', 'BareItem', 'AlbumFile', 'Seq', 'FileError', 'Dest', 'PiwigoId', 'LastPic']

        # Columns used in References.xlsx
        self.rpParent = 4 # E
        self.rpChild = 6 # G
        self.rpSeq = 9 # J
        self.rpLastCol = 10 # K
        
        self.colsRefs = ['rpParent', 'rpChild', 'rpSeq', 'rpLastCol']
        