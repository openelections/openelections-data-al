"""
Converts precinct-level spreadsheet files to CSV

https://github.com/openelections/openelections-data-al/issues/1

1) Download ZIP files from http://sos.alabama.gov/alabama-votes/voter/election-data
2) Unzip and save to data/AL


__author__ Devin Brady
__github__ devinbrady
__twitter__ bradyhunch
"""

import pandas as pd
import xlrd
from pathlib import Path
import sys

def main(): 

    elections = pd.read_csv('alabama_general_precinct_files.csv')

    for elec_name in elections.election_name: 

        if elec_name == '2016 General Election Results - Precinct Level': 
            read_election_directory_2014_or_later(elec_name)


def read_election_directory_2014_or_later(directory):

    p = Path('data/AL/' + directory)
    counties = [f for f in p.iterdir()]

    results_dict = {}

    for c in counties:         

        # a really hacky way of extracting county names
        county_name = c.name
        county_name = county_name.replace('2016-General-', '')
        county_name = county_name.replace('.xlsx', '')
        county_name = county_name.replace('.xls', '')

        print('Processing county: ' + county_name)        

        xl = pd.ExcelFile(c)

        # Read the first sheet
        df = xl.parse(0)

        # Process spreadsheet differently depending on the first column
        if df.columns.values[0] == 'Contest Title': 

            # Unpivot the spreadsheet
            melted = pd.melt(df, id_vars=['Contest Title', 'Party', 'Candidate'], var_name='precinct', value_name='votes')

            melted.rename(columns={
                'Contest Title' : 'office'
                , 'Party' : 'party'
                , 'Candidate' : 'candidate'
                }, inplace=True)

            # Strip trailing spaces from all columns
            for c in melted.columns:
                if melted[c].dtype == 'object':
                    melted[c] = melted[c].str.strip()


            # TODO: need to finish splitting out 'office' and 'district' here

            offices_with_districts = [
                'UNITED STATES REPRESENTATIVE, '
                , 'STATE BOARD OF EDUCATION MEMBER DISTRICT '
            ]

            for office in offices_with_districts: 
                districts = melted.office.str.extract(office + '(\d)', expand=False)
                break


            results_dict[county_name] = melted[['precinct', 'office', 'party', 'candidate', 'votes']]

        else: 

            print('Not yet able to process this county: ' + c)


    # Concat county results into one dataframe, and save to CSV
    results = pd.concat(results_dict).reset_index()
    results.drop('level_1', axis=1, inplace=True)
    results.rename(columns={'level_0' : 'county'}, inplace=True)

    output_filename = directory + '.csv'
    results.to_csv(output_filename, index=False)
    print('Output saved to: ' + output_filename)


if __name__ == '__main__':
    main()
