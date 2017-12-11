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
import numpy as np
import xlrd
from pathlib import Path
import sys
import re


def main():

    elections = pd.read_csv('alabama_general_precinct_files.csv')

    for elec_name in elections.election_name:

        if elec_name == '2016 General Election Results - Precinct Level':
            process_election_directory_2014_or_later(elec_name)


def process_election_directory_2014_or_later(directory):

    year = directory[:4]
    print('Election: ' + directory)

    p = Path('data/AL/' + directory)
    counties = [f for f in p.iterdir()]

    statewide_dict = {}

    for c in counties:

        # a really hacky way of extracting county names
        county_name = c.name
        county_name = county_name[5:].replace('General-', '').replace('.xlsx', '').replace('.xls', '')

        print('County: ' + county_name)

        xl = pd.ExcelFile(c)

        # Read the first sheet
        df = xl.parse(0)

        # Process spreadsheet differently depending on the first column
        if df.columns.values[0] == 'Contest Title':

            # Unpivot the spreadsheet
            melted = pd.melt(df, id_vars=['Contest Title', 'Party', 'Candidate'], var_name='precinct', value_name='votes')

            # Rename columns to match standard
            melted.rename(columns={
                'Party' : 'party'
                , 'Candidate' : 'candidate'
                }, inplace=True)

            # Strip trailing spaces from all columns
            for c in melted.columns:
                if melted[c].dtype == 'object':
                    melted[c] = melted[c].str.strip()

            # A dictionary containing districted office names, and the exact string that precedes each district number
            offices_with_districts = {
                'UNITED STATES REPRESENTATIVE': 'UNITED STATES REPRESENTATIVE, '
                , 'STATE SENATOR' : 'STATE SENATOR, DISTRICT NO. '
                , 'STATE REPRESENTATIVE' : 'STATE REPRESENTATIVE, DISTRICT NO. '
            }

            melted['office'] = melted['Contest Title']
            melted['district'] = np.nan

            for office in offices_with_districts.keys():

                # Rows containing votes for this office
                office_idx = melted[melted['Contest Title'].str.contains(office)].index

                melted.loc[office_idx, 'office'] = office


                # Extract district number from Contest Title

                # Option A: Column-wise
                # This is saving NULLs and I have no idea why
                # melted.loc[office_idx, 'district'] = melted.loc[office_idx, 'Contest Title'].str.extract(offices_with_districts[office] + '(\d)', expand=True)

                # Option B: Row-wise
                # Slower but it works
                for i in office_idx:
                    regex_match = re.search(offices_with_districts[office] + '\d+', melted.loc[i, 'Contest Title'])
                    regex_group = regex_match.group()
                    regex_final = regex_group.replace(offices_with_districts[office], '')

                    melted.loc[i, 'district'] = regex_final


            statewide_dict[county_name] = melted[['precinct', 'office', 'district', 'party', 'candidate', 'votes']]

        else:

            print('Not yet able to process this county: {}'.format(c))


    # Concat county results into one dataframe, and save to CSV
    statewide = pd.concat(statewide_dict).reset_index()
    statewide.drop('level_1', axis=1, inplace=True)
    statewide.rename(columns={'level_0' : 'county'}, inplace=True)

    output_filename = directory + '.csv'
    statewide.to_csv(output_filename, index=False)
    print('Output saved to: ' + output_filename)

    save_presidential_vote_by_county(statewide, year)
    save_us_house_vote_by_district(statewide, year)


def save_presidential_vote_by_county(statewide, year):
    """Checks presidential vote by county

    Save to CSV in order to confirm results.

    Note that this count includes provisional ballots.
    """

    presidential_by_county = pd.pivot_table(
        statewide[statewide.office == 'PRESIDENT AND VICE PRESIDENT OF THE UNITED STATES']
        , columns='candidate'
        , index='county'
        , values='votes'
        , aggfunc='sum'
        , fill_value=0
        , margins = True
        , margins_name = 'Total'
        )

    presidential_by_county.to_csv('{}_presidential_by_county.csv'.format(year))


def save_us_house_vote_by_district(statewide, year):
    """Checks US House vote by district

    Save to CSV in order to confirm results.

    Note that this count includes provisional ballots.
    """

    us_house_districts = pd.pivot_table(
        statewide[statewide.office == 'UNITED STATES REPRESENTATIVE']
        , columns='party'
        , index='district'
        , values='votes'
        , aggfunc='sum'
        , fill_value=0
        , margins = True
        , margins_name = 'Total'
        )

    us_house_districts.to_csv('{}_us_house_districts.csv'.format(year))


if __name__ == '__main__':
    main()
