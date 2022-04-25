# Excel Logger


## Description

Designed to filter through Excel spreadsheets and pull out specific information to add to a log file.

 - Will search through a folder of files, filtering out non-Excel files and moving them to an error folder.
 - Loops through the data of the Excel files that remain and prints out the data relating to the month specified in the file title.
 - Any errored files found in the process will also be moved to the error folder.
 - All files that have been processed successfully are then moved to an archive folder.

## Built With

This application is built with Python and uses several Python libraries, such as datetime and openpyxl

## Installation

1. To install this website to a local computer, clone the repository using the following command

```
git clone git@github.com:MSlanaker/excel-logger.git
```

## Usage

Run the "Python logger.py" command and the program will execute automatically, searching through any files in the call spreadsheets folder.
