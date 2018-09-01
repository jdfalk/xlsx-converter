# XLSX Converter
import argparse
import logging
import os
import sys

import pandas as PD


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-d', '--dir', help='Directory to start in', default=os.getcwd())
    parser.add_argument('-o', '--output', help='File path and name for output')
    parser.add_argument('-e', '--header', help='Header row for file', default="foo,bar")
    parser.add_argument('-s', '--sheet', help='Identify Sheet in spreadsheet')
    parser.add_argument('-l', '--log-level', help='Logging level (default WARNING)',
                        default='WARNING')
    args = parser.parse_args()

    if args.output is None:
        sys.exit("Need output file. See xlsx_converter.py --help")

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

    # converting the path this program receives to
    # what python understands
    output_file = os.path.abspath(args.output)
    header = args.header + "\n"

    # sets the header row of the output file
    # added mode="a"because default mode is r
    # r = read only
    with open(output_file, mode="a") as f:
        f.write(header)

    # looks at a specific directory, and exports the files into a list with the full path
    # after the word "for" it declares
    for root, _, files in os.walk(args.dir):
        for file in files:
            if file.endswith((".xls", ".xlsx")):
                # declares the meaning of variable "full_path" used in this code
                full_path = os.path.join(root, file)
                # adding logging.debug allows the program to print the list if you want
                # to see it when you enable debug logging on
                # the program "full_path is: "
                # is a string that displays the variable "full_path" and
                # the added characters "is: "
                # logging.debug uses c style formatting, the %s defines
                # whatever goes there as a string.
                logging.debug("full_path is: %s", full_path)
                # read excel file into a dataframe, declare details about the dataframe.
                # We assumed all default values.
                try:
                    data_xls = PD.read_excel(
                        io=full_path,
                        sheet_name=int(args.sheet),
                        skiprows=0,
                        header=1
                        )
                    data_xls.to_csv(
                        output_file,
                        mode='a',
                        header=False,
                        index=False
                    )
                except IndexError as err:
                    logging.error("Error occured: %s", err)
                    logging.error("Bad file %s. Skipping and continuing", full_path)
                    continue


if __name__ == '__main__':
    main()
