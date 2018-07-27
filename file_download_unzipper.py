"""
Grabs precinct data zip files from csv containing, election name and url

Downloads them to a data folder for later processing

__author__ Nathan Danielsen
__github__ ndanielsen
__twitter__ nate_somewhere

March 5th 2017 at NICAR2017
"""

import csv
import os
import requests
import zipfile
import glob

def open_files_to_download(csv_filepath):
    """
    Opens a csv file with election name and zipfile location

    Returns a list of dictionaries (key -> election name, value-> datafile_url )
    """
    with open(csv_filepath, 'r') as csvfile:
        data = list(csv.DictReader(csvfile))

    return data

def download_to_folder(filename, file_url, statename='AL'):
    """
    Downloads files to a specific state data folder
    If data dir or state dir do not exist, it creates it for you
    """
    download_destination = os.path.join('data', statename)
    if not os.path.exists(download_destination):
        os.makedirs(download_destination)

    file_path = os.path.join(download_destination, filename)
    print(f"Downloading {file_url}...")
    r = requests.get(file_url)
    with open(file_path, 'wb') as filename:
        filename.write(r.content)

def unzip_zip_files(datadir, destination_path=None):
    """
    Mass unzips all zip files located in your specied data folder

    unzips them into a file folder with the same name as zip file
    """
    if not destination_path:
        destination_path = datadir

    zip_files = glob.glob(f'{datadir}/**/*.zip', recursive=True)

    for zip_file in zip_files:
        zip_file_name = os.path.basename(zip_file)
        folder_name = zip_file_name.replace('.zip', '').lower()
        zip_destination = os.path.join(destination_path, folder_name)

        try:
            with zipfile.ZipFile(zip_file,"r") as zip_ref:
                zip_ref.extractall(zip_destination)
        except:
            print(f"ERROR: Can't unzip {zip_file}")



if __name__ == '__main__':
    csv_file = "alabama_general_precinct_files.csv"
    file_urls = open_files_to_download(csv_file)
    for electionname in file_urls:
        fileurl = electionname['zipurl']
        filename = fileurl.split('/')[-1] # files the filename in the filepath
        filename = filename.replace('.exe', '.zip') # Strangely, the 2004 general is saved as an .exe when it's really a .zip
        download_to_folder(filename, fileurl)

    unzip_zip_files('data/AL')
