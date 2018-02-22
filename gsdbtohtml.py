#!/usr/bin/env python
# -*- coding: utf-8 -*-

#title           :gsdbtohtml.py
#description     :Generate index.html HTML table from Google Sheets dataset
#author          :Mateo Wartelle
#date            :20170222
#usage           :python gsdbtohtml.py
#notes           :
#python_version  :3.6.2
#==============================================================================

from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from oauth2client import file, client, tools
from apiclient import discovery
from httplib2 import Http
import oauth2client
from collections import OrderedDict
import re
import string

def convert_rgb_to_hex(ListTuple):
    """Returns hex when given a triple Tuple of rgb"""
    for (r,g,b) in ListTuple:
        retHEX = '%02x%02x%02x' % (int(r*255), int(g*255), int(b*255))
        return retHEX
    
def html_table(data):
    """Creates a new index.html and writes the list of list into a table"""
    Html_file= open("index.html","w")
    for sublist in data:
        Html_file.write ('<tr><td>')
        Html_file.write ('</td><td>'.join(sublist))
        Html_file.write ('</td></tr>')
    Html_file.write ('</table>')
    Html_file.write("</center>")
    Html_file.close()

def get_credentials():
    """Authenticates with google sheets API returns connection"""
    SCOPES = 'https://www.googleapis.com/auth/drive.readonly'
    store = file.Storage('storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
        creds = tools.run_flow(flow, store)
    DRIVE = discovery.build('sheets', 'v4', http=creds.authorize(Http()))
    return DRIVE

def color_data_extraction(ranges_param):
    """Authenticates and returns list of data and list of RGB from cells"""
    service = get_credentials()
    spreadsheet_id = 'YOUR SPREEDSHEET ID'
    ranges = ranges_param
    num_of_rows, num_of_cols = determine_rows_cols(ranges_param)
    include_grid_data = True
    request = service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=ranges, includeGridData=include_grid_data)
    response = request.execute()
    rowindex = 0
    colindex = 0
    Dataset = []
    RGBDataset = []
    while rowindex <= num_of_rows:
        while colindex <= num_of_cols:
            try:
                data_in_cell = (response['sheets'][0]['data'][0]['rowData'][colindex]['values'][rowindex]['formattedValue'])
            except (KeyError,IndexError):
                data_in_cell = ''
            try:
                keys = (list(response['sheets'][0]['data'][0]['rowData'][colindex]['values'][rowindex]['effectiveFormat']['backgroundColor'].keys()))
                values = (list(response['sheets'][0]['data'][0]['rowData'][colindex]['values'][rowindex]['effectiveFormat']['backgroundColor'].values()))
                mapped = dict(zip(keys, values))
            except (KeyError,IndexError):
                keys = ['red', 'green', 'blue']
                values = [1, 1, 1]
                mapped = dict(zip(keys, values))
            if 'green' not in mapped:
                mapped['green'] = 0
            if 'blue' not in mapped:
                mapped['blue'] = 0
            if 'red' not in mapped:
                mapped['red'] = 0   
            sorted_dictionary = OrderedDict(sorted(mapped.items(), key=lambda v: v, reverse=True))
            rgb_dictionary = (dict(sorted_dictionary))
            rgb_tuple = [tuple(rgb_dictionary.values()) for d in rgb_dictionary]
            color_of_cell = convert_rgb_to_hex(rgb_tuple)
            RGBDataset.append("#"+color_of_cell)
            Dataset.append(data_in_cell)            
            composite_list = [Dataset[x:x+num_of_rows] for x in range(0, len(Dataset),num_of_rows)]
            rowindex = rowindex + 1
            if rowindex >= num_of_rows:
                rowindex = 0
                colindex = colindex + 1
            if colindex == num_of_cols:
                colindex = 0
                break
        break
    return composite_list, RGBDataset

def read_in_index_paint(file, RGBDataset):
    """Reads in file and paints the td cells with the correct color from RGBDataset list"""
    Open_index = open(file,"r")
    test = Open_index.read()
    test = test.split("<td>")
    test = test[1::]
    end = []
    for element, RGB in zip(test, RGBDataset):
        end.append('<td align="center" BGCOLOR="{}">'.format(RGB) + element)
    end_write = "".join(end)
    Html_file= open("index.html","w")
    Html_file.write("""<link rel="stylesheet" href="styles.css">""")
    Html_file.write("""<meta http-equiv="refresh" content="5"/>""")
    Html_file.write ('<table style="height:100%;width:100%; position: absolute; top: 0; bottom: 0; left: 0; right: 0;border:1px solid">')
    Html_file.write(end_write)
    Html_file.write ('</table>')
    Html_file.close()
    Open_index.close()

def determine_rows_cols(range_param):
    """Return num_of_rows, num_of_cols with a given string range"""
     if len(range_param) < 1:
         print("Correct the format of your range (ie 'A1:E3')")
         return
     range_param = range_param.strip()
     range_param = range_param.replace(":","")
     temp = re.split('(\d+)',range_param)
     temp = [x for x in temp if x]
     try:
          Row_data_left = temp[0]
          Row_data_right = temp[2]
          Col_data_left = temp[1]
          Col_data_right = temp[3]
          num_of_rows = convert_letter_to_number(Row_data_right) - convert_letter_to_number(Row_data_left) + 1
          num_of_cols = int(Col_data_right) - int(Col_data_left) + 1
          return num_of_rows, num_of_cols
     except (IndexError):
          print("Correct the format of your range (ie 'A1:E3'")

def convert_letter_to_number(Letter):
    """Converts a letter to the position number in the alphabet"""
     dict_alpha = {c: i for i,c in enumerate(string.ascii_uppercase, 1)}
     if len(Letter) < 1 or len(Letter) > 1:
         print("Correct the letter conversion")
         return
     Letter = Letter.upper()
     value_number = dict_alpha.get(Letter)
     return value_number
    
if __name__ == "__main__":
    """Calls color_data_extraction and returns [List] composite_list and [List] RGBDataset """
    composite_list, RGBDataset = color_data_extraction('A1:D4')
    """Generate html_table from the list of list composite_list"""
    html_table(composite_list)
    """Reads back in the file and paints the cells with the correct bgcolor from RGBDataset"""
    read_in_index_paint("index.html", RGBDataset)