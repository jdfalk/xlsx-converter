# XLSX Converter
import pandas as PD


with open('$HOME/converted.csv') as f:
    f.write("column1,column2,column3\n")



# read excel file into a dataframe, declare details about the dataframe.  We assumed all default values.
data_xls = PD.read_excel("$HOME/testbook.xlsx")
data_xls.to_csv(
    "$HOME/converted.csv", 
    mode='a',
    header=False, 
    index=False
    )
