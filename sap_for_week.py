# create the TableMaker object.
from lib.OutlookToPandas import *
outlook = OutlookToPandas()
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-y', '--year')
    parser.add_argument('-w', '--weeknumber')
    args = parser.parse_args()
    outlook = OutlookToPandas()
    print(outlook.create_week_sap_report(args.year, args.weeknumber))
    print('This is now copied to your clipboard. Paste it in SAP. ')

if __name__ == "__main__":
    main()