#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

# The MIT License (MIT)
# Copyright (c) 2018 OpenElections
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all 
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE 
# SOFTWARE.


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
import os, sys
import re
import glob
import argparse

def main():
    args = parseArguments()

    # for election_dir in glob.glob('data/AL/*/'):
    processor = XLSProcessor(args.inDirPath, args.outFilePath)

    if processor and processor.supported:
        processor.process_election_directory()

def parseArguments():
    parser = argparse.ArgumentParser(description='Parse Alabama vote files into OpenElections format')
    parser.add_argument('inDirPath', type=str,
                        help='path to the Alabama directory given election')
    parser.add_argument('outFilePath', type=str,
                        help='path to output the CSV file to')


    return parser.parse_args()


class XLSProcessor(object):

        # Return the appropriate subclass based on the path
    def __new__(cls, inDirPath, outFilePath):
        if cls is XLSProcessor:
            (dirparent, deepest_dirname) = os.path.split(os.path.dirname(inDirPath))
            m = re.match(r'(20\d\d)', deepest_dirname)
            if m:
                year = int(m.group(1))

            if year >= 2014:
                return super(XLSProcessor, cls).__new__(XLSProcessor_2014)

        else:
            return super(XLSProcessor, cls).__new__(cls, inDirPath, outFilePath)

    def __init__(self, inDirPath, outFilePath):
        self.path = inDirPath
        self.outFilePath = outFilePath
        dirname = os.path.dirname(self.path)
        self.year = dirname[:4]
        self.supported = False
        self.statewide_dict = {}

        self.office_map = {'President And Vice President Of The United States': 'President',
                           'United States Representative': 'U.S. House',
                           'United States Senator': 'U.S. Senate'}
        self.candidate_map = {'Write-In': 'Write-ins'}
        self.valid_offices = frozenset(['President', 'U.S. Senate', 'U.S. House', 'Governor', 'Lieutenant Governor', 'State Senate', 'State House', 'Attorney General', 'Secretary of State', 'State Treasurer',])


    def process_election_directory(self):
        print('Election: ' + self.path)

        for countyFile in glob.glob(f'{self.path}/*'):
            print(countyFile)
            m = re.match(r'\d{4}-(General|Primary)-(.*)\.(csv|xlsx|xls)', os.path.basename(countyFile))

            if m:
                county_name = m.group(2)

                print('==> County: ' + county_name)

                if m.group(3) == 'xlsx' or m.group(3) == 'xls':
                    self.process_excel_file(countyFile, county_name)
                elif m.group(3) == 'csv':
                    self.process_csv_file(countyFile, county_name)
            # break # end after one, for debugging

        # Concat county results into one dataframe, and save to CSV
        statewide = pd.concat(self.statewide_dict).reset_index()
        statewide.drop('level_1', axis=1, inplace=True)
        statewide.rename(columns={'level_0' : 'county'}, inplace=True)

        statewide.to_csv(self.outFilePath, index=False, float_format='%.f')
        print('Output saved to: ' + self.outFilePath)

        # save_presidential_vote_by_county(statewide, year)
        # save_us_house_vote_by_district(statewide, year)

        # print(f"Results for {self.statewide_dict.keys()}")

    def process_excel_file(self, filename, county):
        xl = pd.ExcelFile(filename)

        # Read the first sheet
        df = xl.parse(0)
        # print(df.to_string())

        # Process spreadsheet differently depending on the first column
        if df.columns.values[0] == 'Contest Title':

            # Rename columns to match standard
            df.rename(columns={ 'Party Code': 'party',
                                'Party': 'party',
                                'Candidate': 'candidate',
                                'Candidate Name': 'candidate',
                                }, inplace=True) # Some normalization to do

            # Unpivot the spreadsheet
            melted = pd.melt(df, id_vars=['Contest Title', 'party', 'candidate'], var_name='precinct', value_name='votes')

            # Strip trailing spaces from all columns
            for col in melted.columns:
                if melted[col].dtype == 'object':
                    melted[col] = melted[col].str.strip()

            # Replace empty votes with 0
            melted['votes'] = melted['votes'].fillna(0)

            # A dictionary containing districted office names, and the exact string that precedes each district number
            offices_with_districts = {
                'UNITED STATES REPRESENTATIVE': 'UNITED STATES REPRESENTATIVE, '
                , 'STATE SENATOR' : 'STATE SENATOR, DISTRICT NO. '
                , 'STATE REPRESENTATIVE' : 'STATE REPRESENTATIVE, DISTRICT NO. '
            }

            melted['office'] = melted['Contest Title'] # duplicate the contest title into the office column
            melted['district'] = np.nan

            for office in offices_with_districts:

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
                    regex_match = re.search('(\d+)', melted.loc[i, 'Contest Title'])
                    melted.loc[i, 'district'] = regex_match.group(1)

            melted.office = melted.office.str.title()
            
            # Normalize the office names
            normalize_offices = lambda o: self.office_map[o] if o in self.office_map else o
            melted.office = melted.office.map(normalize_offices)

            # Drop non-statewide offices in place
            mask = melted[~melted.office.isin(self.valid_offices)]
            melted.drop(mask.index, inplace=True)

            # Normalize pseudo-candidates
            normalize_pseudocandidates = lambda c: self.candidate_map[c] if c in self.candidate_map else c
            melted.candidate = melted.candidate.map(normalize_pseudocandidates)

            self.statewide_dict[county] = melted[['precinct', 'office', 'district', 'party', 'candidate', 'votes']]

        else:
            print('Not yet able to process this county: {}'.format(county))


    def process_csv_file(self, filename, county):
        print("process_csv_file not implemented yet")

    def save_presidential_vote_by_county(self, statewide, year):
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


    def save_us_house_vote_by_district(self, statewide, year):
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

class XLSProcessor_2014(XLSProcessor):
    def __init__(self, inDirPath, outFilePath):
        super().__init__(inDirPath, outFilePath)
        self.supported = True ### WRONG…


if __name__ == '__main__':
    main()
