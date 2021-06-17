# xls-convert

VBA module that converts all .xls files in a given directory to .csv. 

Written to process reports from 12 year-old software, which puts out files that are nominally .xls, but which Excel no longer recognizes as valid files.\* The header information suggests they're some sort of XML/HTML table, but I wasn't able to get anything to convert or even read them as html files either.

(Also tried using various Python libraries including pandas.read_excel and .read_html, xlrd/xlwt, even openpyxl; none of them were able to read these files either. So Excel VBA was more or less my only option for converting them.)

Anyway! This module will prompt the user to select a source directory (and/or save directory, see option 3 below). When selected, it will automatically export all .xls files in this directory to .csv files, either in the same directory or in another directory of the user's choosing.

This script has to be run from within Excel, usually by importing the .bas as a module. When it imports, you'll see there are several configurable options at the top, including: 1.) an option to delete the first *x* rows (these are an unneccesary pseudo-header in the files I was working from); 2.) whether or not to prompt the user to confirm their folder selections; 3.) an option to either save to the same folder as the source file or select a different save folder; and 4.) an option to keep the same filename or create a custom filename using the contents of the .xls file.



\* When opening one of these weird .xls files, Excel throws the error message: 

>"The file format and extension of _filename_ don't match. The file could be corrupted or unsafe. Unless you trust its source, don't open it. Do you want to open it anyway?"

I include this error message to help out anyone else who's Googling the same issue! Definitely try out the Python libraries I mentioned above first, but if you aren't getting anywhere then feel free to use this script instead. Just keep in mind that it is customized to fit the filenames and file structure of the .xls files I had to work with, so some things will need to be changed - the custom filename option comes to mind as being basically unusable outside my specific use case. :)
