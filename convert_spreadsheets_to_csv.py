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

    if processor: # and processor.supported:
        processor.process_election_directory()

def parseArguments():
    parser = argparse.ArgumentParser(description='Parse Alabama vote files into OpenElections format')
    parser.add_argument('inDirPath', type=str,
                        help='path to the Alabama directory given election')
    parser.add_argument('outFilePath', type=str,
                        help='path to output the CSV file to')


    return parser.parse_args()


class XLSProcessor(object):
    def __init__(self, inDirPath, outFilePath):
        self.path = inDirPath
        self.outFilePath = outFilePath
        (dirparent, deepest_dirname) = os.path.split(os.path.dirname(inDirPath))
        self.year = deepest_dirname
        self.supported = False
        self.statewide_dict = {}
        self.completeColumnNames = ['precinct', 'office', 'district', 'party', 'candidate', 'votes']

        self.office_map = {
            'Registered Voters - Total': 'Registered Voters',
            'President And Vice President Of The United States': 'President',
            'President And Vice-President Of The United States': 'President',
            'President Of The United States': 'President',
            'United States Representative': 'U.S. House',
            'US Rep': 'U.S. House',
            'United States Senator': 'U.S. Senate',
            'Governor': 'Governor',
            'Lt. Governor': 'Lieutenant Governor',
            'Attorney General': 'Attorney General',
            'State Treasurer': 'State Treasurer',
            'Commissioner Of Agriculture And Industries': 'Commissioner of Agriculture and Industries',
            'Secretary Of State': 'Secretary of State',
            'State Auditor': 'State Auditor',
            'State Senator': 'State Senate',
            #'State Sen': 'State Senate',
            'State Representative': 'State House',
            #'State Rep': 'State House',
            'Ballots Cast - Total': 'Ballots Cast',
            'STRAIGHT PARTY': 'Straight Party'
            }
        self.candidate_map = {'Write-In': 'Write-ins', 'Write-in': 'Write-ins'}
        self.valid_offices = frozenset(['Registered Voters', 'Ballots Cast', 'Straight Party', 'President', 'U.S. Senate', 'U.S. House', 'Governor', 'Lieutenant Governor', 'Attorney General', 'State Treasurer', 'Commissioner of Agriculture and Industries', 'State Senate', 'State House', 'Secretary of State', 'State Auditor'])


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
            else:
                (county_name, ext) = os.path.basename(countyFile).split(os.extsep, 1)
                self.process_excel_file(countyFile, county_name)
            # break # end after one, for debugging

        # Concat county results into one dataframe, and save to CSV
        statewide = pd.concat(self.statewide_dict).reset_index()
        statewide.drop('level_1', axis=1, inplace=True)
        statewide.rename(columns={'level_0' : 'county'}, inplace=True)

        statewide = statewide.sort_values(['county', 'precinct', 'office', 'district', 'party', 'candidate'], ascending=True)
        statewide.to_csv(self.outFilePath, index=False, float_format='%.f')
        print('Output saved to: ' + self.outFilePath)

        # save_presidential_vote_by_county(statewide, year)
        # save_us_house_vote_by_district(statewide, year)

        # print(f"Results for {self.statewide_dict.keys()}")


    def process_excel_file(self, filename, county):
        xl = pd.ExcelFile(filename)

        # Read the first sheet
        df = xl.parse(0, header=None) # Leave out headers because the two formats use them differently
        df = self.stripCellsDropEmptyRows(df)

        # Process spreadsheet differently depending on the first cell
        firstCell = df.iloc[0, 0] # Contents of very first cell

        # print(f"{firstCell}")
        if firstCell == 'Contest Title':
            self.process_contest_title_excel_file(df, county)
        elif firstCell == 'Table of Contents':
            self.process_TOC_excel_file(xl, df, county)
        elif pd.isnull(firstCell) or firstCell in self.valid_offices:
            self.process_blank_header_excel_file(df, county)
        else:
            print('Not yet able to process this county: {}'.format(county))


    #
    # This format is used in all 2016 xls files, and some of the earlier Excel files
    #
    def process_contest_title_excel_file(self, df, county):
        # return # Temp while writing the alternative branch

        # Set header
        df.columns = df.iloc[0] # set the columns to the first row
        df.reindex(df.index.drop(0)) # reindex, dropping the now-duplicated first row

        # Rename columns to match standard
        df.rename(columns={ 'Party Code': 'party',
                            'Party': 'party',
                            'Candidate': 'candidate',
                            'Candidate Name': 'candidate',
                            }, inplace=True) # Some normalization to do

        # Unpivot the spreadsheet
        melted = pd.melt(df, id_vars=['Contest Title', 'party', 'candidate'], var_name='precinct', value_name='votes')

        melted.dropna(how='any', subset=['votes'], inplace=True) # Drop rows with na for votes

        melted = self.populateOfficesAndDistricts(melted)
        melted = self.normalizeOfficesAndCandidates(melted)

        self.statewide_dict[county] = melted[self.completeColumnNames]

    #
    # This format is used in many 2014 files
    #
    def process_blank_header_excel_file(self, df, county):
        # return # Temp while developing

        # Forward-Fill offices to the right
        df.loc[[0]] = df.loc[[0]].ffill(axis=1)

        # Remove bogus data in Clay 2014 :-(
        if county == 'Clay' and self.year == '2014':
            df = df.drop(df.index[[21, 22, 23]]) # BARF!

        # Transpose
        df = df.transpose()

        # Fix column headers
        df.iat[0, 0] = 'office'
        df.iat[0, 1] = 'candidate'

        # Set header
        df.columns = df.iloc[0] # set the columns to the first row
        df.drop([0], inplace=True)

        # Melt the spreadsheet into an OE-friendly format
        # print(df.iloc[:, 0].head(5))
        melted = pd.melt(df, id_vars=['office', 'candidate'], var_name='precinct', value_name='votes')
        # print("melted")

        melted.dropna(how='any', subset=['votes'], inplace=True) # Drop rows with na for votes

        # Create the new district column and fill with NaN
        melted['district'] = np.nan
        melted['party'] = ''

        # Split out district names from offices
        contests = melted["office"].drop_duplicates()

        for contest in contests:
            # print(f"--- {contest}")
            office, district = self.identifyOfficeAndDistrict(contest)

            if district:
                melted.loc[melted['office'] == contest, 'district'] = district
                melted.loc[melted['office'] == contest, 'office'] = office
                # print(melted.loc[melted['office'] == office])

        # Split out party names from candidates
        candidates = melted["candidate"].drop_duplicates()

        for origCandidate in candidates:
            # print(f"--- {origCandidate}")
            candidate, party = self.identifyCandidateAndParty(origCandidate)
            # print(f"{candidate})

            if party:
                melted.loc[melted['candidate'] == origCandidate, 'party'] = party
                melted.loc[melted['candidate'] == origCandidate, 'candidate'] = candidate
                # print(melted.loc[melted['candidate'] == candidate])

        # Normalize name of "Total" pseudo-precinct
        melted.loc[melted["precinct"] == 'REPORTED TOTALS', 'precinct'] = 'Total'

        # Drop any "CALCULATED TOTALS", which we can recalculate ourselves
        melted = melted.drop(melted[melted["precinct"] == 'CALCULATED TOTALS'].index)

        melted = self.normalizeOfficesAndCandidates(melted)

        self.statewide_dict[county] = melted[self.completeColumnNames]


    def process_TOC_excel_file(self, xl, firstSheetDF, county):
        countyDF = pd.DataFrame()
        for sheetName in self.relevant_sheets(firstSheetDF):
            print(f"--> parse {county} sheet {sheetName}")
            df = xl.parse(sheetName, header=None) # Leave out headers to define our own later
            df = self.stripCellsDropEmptyRows(df)

            # Drop duplicated office
            office = df.iloc[0, 0]
            m = re.compile("(FOR )?([\w, -]+) \(Vote For 1\)").search(office)
            if m:
                office = m.group(2)

            df.drop([0], inplace=True)

            # Ignore superfluous "total" data
            results = df.iloc[:, :-1:2]

            # Fix naming of columns and totals
            results.iat[0, 0] = 'precinct'
            results.iat[-1, 0] = 'Total'

            # Set header
            results.columns = results.iloc[0, :]
            results = results[2:] # Drop the first two rows

            melted = pd.melt(results, id_vars=['precinct'], var_name='candidate', value_name='votes')
            melted['Contest Title'] = office
            melted['party'] = ''
            # import pdb; pdb.set_trace()
            melted = self.populateOfficesAndDistricts(melted)
            melted = self.normalizeOfficesAndCandidates(melted)

            countyDF = countyDF.append(melted[self.completeColumnNames])

        self.statewide_dict[county] = countyDF

    def relevant_sheets(self, df):
        relevantSheetNames = []

        toc = df.iloc[0:, 0:2]
        toc = toc.fillna('')
        relevantOffices = ("FOR PRESIDENT AND VICE", "FOR UNITED STATES REPRESENTATIVE")

        for index, value in toc.iterrows():
            if value[1].startswith(relevantOffices):
                relevantSheetNames.append(str(value[0]))

        return relevantSheetNames

    def process_csv_file(self, filename, county):
        df = pd.read_csv(filename)

        # Normalize column names
        colNames = ['county', 'election_date', 'contest_number', 'candidate_number', 'votes', 'party', 'Contest Title', 'candidate', 'precinct', 'district_name']
        df.columns = colNames

        df = self.stripCellsDropEmptyRows(df)

        # Drop "registered voters" and "ballots cast"
        df = df.loc[df["contest_number"] >= 100]

        df = self.populateOfficesAndDistricts(df)
        df = self.normalizeOfficesAndCandidates(df)

        self.statewide_dict[county] = df[self.completeColumnNames]


    def populateOfficesAndDistricts(self, df):
        df['office'] = df['Contest Title'] # duplicate the contest title into the office column
        df['district'] = np.nan

        # List of all contests
        contests = df["office"].drop_duplicates()

        # Split out district names from offices
        for contest in contests:
            m = re.compile(r'[ ,] (DISTRICT )?(\d+)').search(contest)

            if m:
                # Set district to found number, trim office
                df.loc[df['office'] == contest, 'district'] = m.group(2)
                df.loc[df['office'] == contest, 'office'] = contest[:m.span()[0]] # Strip district number off contest

        return df

    def normalizeOfficesAndCandidates(self, df):
        df.office = df.office.str.title()

        # Normalize the office names
        normalize_offices = lambda o: self.office_map[o] if o in self.office_map else o
        df.office = df.office.map(normalize_offices)

        # Drop non-statewide offices in place
        mask = df[~df.office.isin(self.valid_offices)]
        df.drop(mask.index, inplace=True)

        # Normalize pseudo-candidates
        normalize_pseudocandidates = lambda c: self.candidate_map[c] if c in self.candidate_map else c
        df.candidate = df.candidate.map(normalize_pseudocandidates)

        return df

    def identifyOfficeAndDistrict(self, contest):
        (office, district) = (contest, None)

        try:
            office_district = re.compile('[\W]+[Dd]ist[\W]+').split(contest)

            if len(office_district) > 1:
                office, district = office_district

            recognizedOffice = (office in self.office_map)

            if recognizedOffice:
                office = self.office_map[office]
        except:
            print(f"Couldn't split contest '{contest}'")

        return (office, district)

    def identifyCandidateAndParty(self, origCandidate):
        (candidate, party) = (origCandidate, None)

        if not pd.isnull(origCandidate):
            m = re.compile('\s*\(\s*([\w\.]+)\s*\)\s*').search(origCandidate)

            if m:
                party = m.group(1)
                candidate = origCandidate.replace(m.group(0), '')

        return (candidate, party)

    # Clean the data
    def stripCellsDropEmptyRows(self, df):
        df = df.applymap(lambda x: x.strip() if type(x) is str else x) # Strip all cells
        df = df.replace('', np.NaN, regex=True) # Replace empty cells with NaN
        df = df.dropna(how='all') # Drop rows that only consist of NaN data

        return df

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


if __name__ == '__main__':
    main()
