# Directory Content

This is a collection of scripts for photo album migrations as described under https://piwigo.miraheze.org/wiki/Main_Page

File discriptios as follows:

# addRefs.py

Add Sequence numbering to Album and Photos from input files Pic.xslx, Album.xlsx and Refs.xlsx


# checkPics.py

Perform QC check on photos metadata. It will test the following:

- There is a thumbnail named doc_pic.jpg
- The files are of a reasonable size
- Legal file types, jpeg, jpg, png and gif

# cols.py

Description of coloumns in Album.xlsx and Pic.xlsx as well as some other common variables in the migration scripts

# getPiwigId.py

Script to be executed after Piwigo have digested to map Piwigo ID with source ID. Required to be able to know which photo should have what metadata

# x-port.py

Transforming script. Copying files from the source structure to a Piwigo structure. Also create some metadata

Read an two input files, Album.xlsx and Pic.xlsx for Albums and Photos metadata.
