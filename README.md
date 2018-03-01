## Google Sheets Database to HTML Table
Python Program to generate HTML Table from Google Sheets Database

## Project Overview:

## Solution:
A constant update stream from Google Sheets API displayed in a HTML Table. 

## Example:
![Screenshot](Example.png)

##  How to install & Use<br>
`git clone https://github.com/MateoWartelle/GoogleSheetsDB-to-HTML.git`<br>
`cd gsDB-to-HTML`<br>
`Edit 'YOUR SPREEDSHEET ID' in color_data_extraction function`<br>
`Set the proper range in main (ie 'A1:E5')`<br>
`run python gsdbtohtml.py`

## Current Problems
- Does not take into account rows and or columns that are merged. 
Currently this can be done manually by adding rowspan=x and colspan = y to each td entry

