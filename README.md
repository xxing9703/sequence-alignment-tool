# Sequence-alignment-tool
 
## Introduction
This is an automated tool for sequence alignment.  For multiple protein sequences of high similarities and comparable length, this tool performs multi-alignment and generates the gene sequences based on the selected condon table. The result is exported to an excel table with highlighted colors showing the identical(green), differed(red) and inserted (yellow) on each aa site in sheet 1 and the resulting gene sequences in sheet 2.
![image](https://user-images.githubusercontent.com/16364863/143302089-2503d519-4826-43f6-ade3-8a43bd7e6928.png)

![image](https://user-images.githubusercontent.com/16364863/143302819-626e3b4b-be6d-4350-bd78-58415107fff9.png)

## Run an example
required input file:
1. A condon table in excel file (condon.xlsx) of user's choice, which is editable and can contain up to 3 different ones in columns. 
2. An input excel file containing mutliple protein sequences (names in column 1, sequences in column 2) and the gene sequence of the first one only in column 3. some examples are included. 
To start, run "gui_protein_seq.m" in matlab. 1) select a condon, 2) load an input file, 3) press run  4) save to excel
![image](https://user-images.githubusercontent.com/16364863/143302169-763b179e-cbdb-42fa-94d8-39cb6aa962fc.png)

## Requirments
Installation of Matlab R2018b or higher with bioinformatic tools box is required.
A deployable version without the need of a matlab licsence is available upon request.
