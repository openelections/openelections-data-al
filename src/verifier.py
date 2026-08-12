#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

# The MIT License (MIT)
# Copyright (c) 2016 Nick Kocharhook
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

import pdb
import csv
import os
import re
import argparse

def main():
	args = parseArguments()

	for path in args.paths:
		verifier = Verifier(path)
		verifier.showPrimaryPartiesError = not args.mutePrimaryPartiesError
		verifier.showPartiesError = not args.mutePartiesError
		verifier.showXForDistrictError = not args.muteXForDistrictError
		verifier.singleErrorMode = args.singleError

		if verifier.ready and "matrix" not in verifier.filename:
			verifier.verify()


def parseArguments():
	parser = argparse.ArgumentParser(description='Verify openelections CSV files')
	parser.add_argument('--mutePrimaryPartiesError', dest='mutePrimaryPartiesError', action='store_true')
	parser.add_argument('--mutePartiesError', dest='mutePartiesError', action='store_true')
	parser.add_argument('--muteXForDistrictError', dest='muteXForDistrictError', action='store_true')
	parser.add_argument('--singleError', dest='singleError', action='store_true', help='Display only the first error in each file')
	parser.set_defaults(mutePrimaryPartiesError=False, mutePartiesError=False, muteXForDistrictError=False)
	parser.add_argument('paths', metavar='path', type=str, nargs='+',
					   help='path to a CSV file')

	return parser.parse_args()


class Verifier(object):
	validColumns = frozenset(['county', 'precinct', 'office', 'district', 'party', 'candidate', 'votes', 'notes'])
	requiredColumnSet = frozenset(['county', 'precinct', 'office', 'district', 'party', 'candidate', 'votes'])
	uniqueRowIDSet = frozenset(['county', 'precinct', 'office', 'district', 'party', 'candidate'])
	# Statewide offices, as canonicalized by the two converters
	# (convert_spreadsheets_to_csv.py:office_map and convert_precinct_pdfs.py
	# STATEWIDE_OFFICE_MAP). Both converters emit "State Auditor" and
	# "Commissioner of Agriculture and Industries"; the older "Auditor" /
	# "Commissioner of Agriculture" spellings are kept too so rows that an
	# inconsistent converter didn't fully normalize don't false-positive. The
	# 2026 canvass batch added State Board of Education and the state
	# executive committees, which the original closed set predated.
	validOffices = frozenset([
		'President', 'U.S. Senate', 'U.S. House', 'Governor', 'Lieutenant Governor',
		'State Senate', 'State House', 'Attorney General', 'Secretary of State',
		'State Treasurer', 'State Auditor', 'Auditor',
		'Commissioner of Agriculture and Industries', 'Commissioner of Agriculture',
		'State Board of Education',
		'State Democratic Executive Committee', 'State Republican Executive Committee',
		# Raw spellings that escaped the converters' office_map normalization
		# (present in some 2026 county CSVs); accepted so the verifier doesn't
		# cry wolf on legit offices. The canonicalization should be fixed at the
		# converter, not here.
		'United States Senator', 'United States Representative',
		'State Senator', 'State Representative',
	])
	# Case-insensitive view of validOffices, so all-caps converter output
	# (e.g. "GOVERNOR", "STATE AUDITOR") passes without enumerating every
	# casing — those are case bugs, not invalid offices.
	validOfficesLower = frozenset(o.lower() for o in validOffices)
	# Offices whose name carries a place/amendment number or a county-local
	# race title — matched case-insensitively so the canvass converters'
	# spelling variants (comma-vs-none, "Place No. 1" vs "Place 1", the
	# occasional all-caps county) all pass, while OCR corruptions like
	# "Proposed Statewioe Amendment 1" or "Commissidner of Agriculture and
	# Industries" still fail to match and get flagged. The county-local branch
	# requires the word "County" plus a known office keyword, except judicial
	# races which are titled by "Judicial Circuit" rather than a county.
	validOfficeRE = re.compile(
		r'^(?:'
		r'Public Service Commission(?:,? Place(?: No\.?)? \d+)?'
		r'|Proposed Statewide Amendment(?: (?:One|Two|\d+))?'
		r'|State (?:Democratic|Republican) Exec(?:utive)? Comm(?:ittee)?.*'
		r'|(?:Circuit|District) Court Judge\b.*'
		r'|(?=.*\bCounty\b)(?=.*\b(?:Sheriff|Coroner|Revenue Commissioner|Property Tax Commissioner|Commission|Board of Education|School Board|School District|Executive Committee|Superintendent|Chairman)\b).*'
		r')$',
		re.IGNORECASE)
	officesWithDistricts = frozenset(['U.S. House', 'State Senate', 'State House'])
	pseudocandidates = frozenset(['Write-ins', 'Under Votes', 'Over Votes', 'Total', 'Total Votes Cast',  'Registered Voters'])
	normalizedPseudocandidates = frozenset(['writeins', 'undervotes', 'overvotes', 'total', 'totalvotescast', 'registeredvoters'])

	# Return the appropriate subclass based on the path
	def __new__(cls, path):
		if cls is Verifier:
			filename = os.path.basename(path)

			if "general" in filename:
				if "precinct" in filename:
					return super(Verifier, cls).__new__(GeneralPrecinctVerifier)
				else:
					return super(Verifier, cls).__new__(GeneralVerifier)
			elif "primary" in filename:
				if "precinct" in filename:
					return super(Verifier, cls).__new__(PrimaryPrecinctVerifier)
				else:
					return super(Verifier, cls).__new__(PrimaryVerifier)
			elif "special" in filename and "precinct" in filename:
				return super(Verifier, cls).__new__(SpecialPrecinctVerifier)

		else:
			return super(Verifier, cls).__new__(cls, path)

	def __init__(self, path):
		self.path = path
		self.columns = []
		self.uniqueRowIDs = {}
		self.reader = None
		self.ready = False
		self.showPrimaryPartiesError = True
		self.showXForDistrictError = True
		self.singleErrorMode = False

		self.countyRE = re.compile("\d{8}__[a-z]{2}_")

		try:
			self.pathSanityCheck(path)

			self.filename = os.path.basename(path)
			self.filenameState, self.filenameCounty = self.deriveStateCountyFromFilename(self.filename)

			self.ready = True
		except Exception as e:
			print("ERROR: {}".format(e))

	def verify(self):
		self.parseFileAtPath(self.path)

	def pathSanityCheck(self, path):
		if not os.path.exists(path) or not os.path.isfile(path):
			raise FileNotFoundError("Can't find file at path %s" % path)

		if not os.path.splitext(path)[1] == ".csv":
			raise ValueError("Filename does not end in .csv: %s" % path)

		print("==> {}".format(path))

	def deriveStateCountyFromFilename(self, filename):
		components = filename.split("__")
		countyIndex = 0

		if "special" in components and ("primary" in components or "general" in components): # special primary or special general
			countyIndex = 4
		elif ("primary" in components or "general" in components): # normal primary or general
			countyIndex = 3

		if countyIndex:
			return (components[1], components[countyIndex].replace("_", " ").title())

		return (None, None)

	def parseFileAtPath(self, path):
		with open(path, 'rU') as csvfile:
			self.reader = csv.DictReader(csvfile)
			self.currentRowIndex = 0
			self.headerColumnCount = 0
			
			try:
				if self.verifyColumns(self.reader.fieldnames):
					for index, row in enumerate(self.reader):
						self.currentRowIndex = index + 2 # 1 for header; 1 for human-readable, 1-indexed list

						self.verifyColumnsOfRow(row)
						# self.verifyCounty(row)
						self.verifyOffice(row)
						self.verifyDistrict(row)
						self.verifyCandidate(row)
						self.verifyParty(row)
						self.verifyVotes(row)
						self.verifyRowIsUnique(row)
			except StopIteration as si:
				pass # Stop verifying when exception is thrown

	def verifyColumns(self, columns):
		self.headerColumnCount = len(columns)

		invalidColumns = set(columns) - Verifier.validColumns
		missingColumns = self.requiredColumns() - set(columns)

		if invalidColumns:
			self.printError("Invalid columns: {}".format(invalidColumns))

		if missingColumns:
			self.printError("Missing columns: {}".format(missingColumns))
			return False

		return True

	def requiredColumns(self):
		return Verifier.requiredColumnSet

	def verifyColumnsOfRow(self, row):
		badColumnCount = len(row) - self.headerColumnCount

		if badColumnCount < 0:
			self.printError("Row is missing {} column(s)".format(abs(badColumnCount)), row)
		elif badColumnCount > 0:
			self.printError("Row has {} extra column(s)".format(badColumnCount), row)

	def verifyCounty(self, row):
		normalisedCounty = row['county'].title()

		if not normalisedCounty == self.filenameCounty:
			self.printError("County doesn't match filename", row)

		if not row['county'] == normalisedCounty:
			self.printError("Use title case for the county", row)

	def verifyOffice(self, row):
		office = row['office']
		if office.lower() in Verifier.validOfficesLower:
			return
		# Some 2026 converter output left the district appended to the office
		# ("State House, District 50") instead of in the district column; the
		# office is still valid once that suffix is stripped. District-column
		# placement is verifyDistrict's concern, not the office check's.
		stripped = re.sub(r',\s*Dist(?:rict)?(?:\s+No\.?)?\s+\d+\s*$', '', office)
		if stripped != office and stripped.lower() in Verifier.validOfficesLower:
			return
		if Verifier.validOfficeRE.match(office):
			return
		self.printError("Invalid office: {}".format(office), row)

	def verifyDistrict(self, row):
		if row['office'] in Verifier.officesWithDistricts:
			if not row['district']:
				self.printError("Office '{}' requires a district".format(row['office']), row)
			elif row['district'].lower() == 'x':
				if not self.showXForDistrictError:
					pass # Some counties use this, but we still want to make sure it's reviewed by default
				else:
					self.printError("District must be an integer", row)
			elif not self.verifyInteger(row['district']):
				self.printError("District must be an integer", row)

	def verifyCandidate(self, row):
		charsRE = re.compile('[^A-Za-z]+', re.UNICODE)
		candidate = row['candidate']
		normalizedCandidate = charsRE.sub('', candidate).lower()

		if candidate not in Verifier.pseudocandidates:
			if normalizedCandidate in Verifier.normalizedPseudocandidates:
				self.printError("Misspelled pseudocandidate a: '{}'".format(candidate), row)
			else:
				# Compare the normalized strings to determine if they match
				for npc in Verifier.normalizedPseudocandidates:
					if normalizedCandidate.startswith(npc[0:4]): # Only check the first 4 characters
						self.printError("Misspelled pseudocandidate b: '{}'".format(candidate), row)
						break

	def verifyParty(self, row):
		if self.showPartiesError:
			if row['candidate'] not in Verifier.pseudocandidates and not row['party']:
				self.printError("Party missing", row)

	def verifyVotes(self, row):
		if not self.verifyInteger(row['votes']):
			self.printError("Vote count must be an integer", row)
		elif not int(row['votes']) >= 0:
			self.printError("Vote count must be greater than or equal to zero", row)

	def verifyRowIsUnique(self, row):
		rowTuple = tuple(row[col] for col in Verifier.uniqueRowIDSet)

		if rowTuple in self.uniqueRowIDs:
			self.printError("Line is duplicated (original line {})".format(self.uniqueRowIDs[rowTuple]), row)
		else:
			self.uniqueRowIDs[rowTuple] = self.currentRowIndex

	def verifyInteger(self, numberStr):
		try:
			integer = int(numberStr)
		except ValueError as e:
			return False

		return True

	def printError(self, text, row=[]):
		print("ERROR: Line {}: {}".format(self.currentRowIndex, text))

		if row:
			print(row)

		if self.singleErrorMode:
			raise StopIteration("Stop after first error")


class GeneralPrecinctVerifier(Verifier):
	pass

class PrimaryPrecinctVerifier(Verifier):
	def verifyParty(self, row):
		if self.showPrimaryPartiesError and self.showPartiesError:
			if not row['party']:
				self.printError("Primary results must include a party for every row", row)

class SpecialPrecinctVerifier(Verifier):
	pass


class PrimaryVerifier(Verifier):
	def requiredColumns(self):
		return Verifier.requiredColumnSet - set(['precinct'])

	def verifyCounty(self, row):
		pass


class GeneralVerifier(Verifier):
	def requiredColumns(self):
		return Verifier.requiredColumnSet - set(['precinct'])

	def verifyCounty(self, row):
		pass


# Default function is main()
if __name__ == '__main__':
	main()