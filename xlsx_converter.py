# XLSX Converter
import logging
import os

import pandas as PD

source_file = "c:\\temp\\Book1.xlsx"
target_file = "c:\\temp\\Book1.csv"

# sets the header row of the output file
with open(target_file) as f:
    f.write("column1,column2,column3\n")

# looks at a specific directory, and exports the files into a list with the full path
# after the word "for" it declares
for root, _, files in os.walk("c:\\temp"):
    for file in files:
        # declares the meaning of variable "full_path" used in this code
        full_path = os.path.join(root, file)
# adding logging.debug allows the program to print the list if you want
# to see it when you enable debug logging on the program
# "full_path is: " is a string that displays the variable "full_path"
# and the added characters "is: "
        logging.debug("full_path is: " + str(full_path))
        # read excel file into a dataframe, declare details about the dataframe.
        # We assumed all default values.
        data_xls = PD.read_excel(full_path)
        data_xls.to_csv(
            target_file,
            mode='a',
            header=False,
            index=False
            )
