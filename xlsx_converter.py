# XLSX Converter
import argparse
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
        # to see it when you enable debug logging on the program "full_path is: " 
        # is a string that displays the variable "full_path" and the added characters "is: "
        # logging.debug uses c style formatting, the %s defines whatever goes there as a string. 
        logging.debug("full_path is: %s", full_path)
        # read excel file into a dataframe, declare details about the dataframe.
        # We assumed all default values.
        data_xls = PD.read_excel(full_path)
        data_xls.to_csv(
            target_file,
            mode='a',
            header=False,
            index=False
            )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-d', '--dir', help='Directory to start in', default=os.getcwd())
    parser.add_argument('-l', '--log-level', help='Logging level (default WARNING)',
                        default='WARNING')
    args = parser.parse_args()
    
    # assuming loglevel is bound to the string value obtained from the
    # command line argument. Convert to upper case to allow the user to
    # specify --log=DEBUG or --log=debug
    numeric_level = getattr(logging, args.log_level.upper(), None)
    if not isinstance(numeric_level, int):
        raise ValueError('Invalid log level: %s' % args.log_level)
    logging.basicConfig(
        level=numeric_level,
        format='%(asctime)s %(levelname)s:%(message)s',
        datefmt='%m/%d/%Y %I:%M:%S %p')

if __name__ == '__main__':
    main()