# Imports
import argparse
import os
import pandas as pd

# Startup message
print("""
   ___ _____   __  _         ___            _ 
  / __/ __\ \ / / | |_ ___  | __|_ ____ ___| |
 | (__\__ \\\\ V /  |  _/ _ \ | _|\ \ / _/ -_) |
  \___|___/ \_/    \__\___/ |___/_\_\__\___|_|
      
            © 2023 Volker Lieber                                       
""")

# Arguments
parser = argparse.ArgumentParser()
parser.add_argument('-i', '--in', type=str, required=True, help='Input folder')
parser.add_argument('-o', '--out', type=str, required=True, help='Output file')
parser.add_argument('-s', '--separator', type=str, help='CSV separator (default ";")', default=';')
args = vars(parser.parse_args())

# Open output file
with pd.ExcelWriter(args['out']) as writer: # pylint: disable=abstract-class-instantiated

    # Check if input folder exists
    if os.path.exists(args['in']):

        # Loop through files in input folder
        for file in os.listdir(args['in']):

            # Check if file is a csv file
            if file.endswith(".csv"):

                # Convert csv to excel
                print(os.path.join(args['in'],file))
                csv = pd.read_csv(os.path.join(args['in'],file), sep=args['separator'])
                csv.to_excel(writer, sheet_name=file.split('.')[0], index = None, header=True)
            else:
                print("No csv files found")
    else:
        print(f"Folder '{args['in']}' does not exist")

print("Done")
