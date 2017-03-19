# GeneCompare
## Purpose 
This program is intended to compare two genearray files. It extracts the log fold change for all the individual genes from the two files you are compairing, it obtains the variace and delivers an excel file containing the genes, various methods of sorting them, and highlights variance for manual interpretation. 

## Required Modules
 - csv - Used to open tab delimited "MTA" files
 - math - Used to calculate variance 
 - openpyxl - Used to create excel spreadsheet output
 - argparse - Used for command line arguments

## Sample Usage: 
 - ./genecompare.py -w whole.tumor.file.txt -c tumor.cells.only.file  -o excel.formatted.output.filename

## Sample Genearray Files: 
 - whole tumor file: "MTA-1_0.norm.gene.matrix.refseq.BPIvsBP.limma.fdr0.01_fc1.5_whole tumor.txt"
 - tumor cells only: "MTA-1_0.norm.gene.matrix.refseq.BPIvsBP.limma.txt.flt.rawp0.01_fc1.5_tumor cells only.txt"
 
## What's inside: 
 - README.md  
 - genecompare.py: Main program
 - getcolors.py: Function to get a color value to be used in formatting the excel file
 - headerizer.py: Function to set formatting of row headers 
