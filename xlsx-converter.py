# XLSX Converter
import pandas as PD

# read excel file into a dataframe, declare details about the dataframe.  We assumed all default values.
data_xls = PD.read_excel("c:\\temp\\Book1.xlsx")
data_xls.to_csv(
    "c:\\temp\\Book1.csv", 
    mode='a', 
    index=False
    )
