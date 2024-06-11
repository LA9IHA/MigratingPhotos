This is a collection of scripts for photo album migrations as described under https://piwigo.miraheze.org/wiki/Main_Page

File discriptios as follows:

# XportHighStage.py

Read an two input files, Album.xlsx and Pic.xlsx for Albums and Photos metadata.

# CheckPicsHighStage.py

Perform QC check on photos metadata. It will test the following:

- There is a thumbnail named doc_pic.jpg
- The files are of a reasonable size
- Legal file types, jpeg, jpg, png and gif

# addSeqHighStage.py

Add Sequence numbering to Album and Photos from input files Pic.xslx, Album.xlsx and Refs.xlsx
