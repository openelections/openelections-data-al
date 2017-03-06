# openelections-data-al
Pre-processed results for Alabama elections

### Getting Started

Note: Developed in python3

Create a python [virtualenv](http://docs.python-guide.org/en/latest/dev/virtualenvs/)

Install the required dependencies
`pip install -r requirements.txt`


Run the downloader and unzipper

`python file_download_unzipper.py`

It will create a data folder with a subfolder called 'AL' for Alabama, download the zip files into that folder and unzip them in similarly named folders in that location.


### Want to Add More Alabama Zipped files?
Add them to `alabama_general_precinct_files.csv` with the name of the election and the zip file location


#### To Do
Extract, wrangle and combine each into a single csv file for each election
