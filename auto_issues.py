__author__ = "Thomas Antonacci"


"""
This script relies on the jiraq bash script and ~awollam/aw/jiraq_parser.py woids.stats.txt
Meant to be run on gsub
"""

#TODO:
#Write column check for new headers from jiraq

import smartsheet
import csv
import os
import sys
import glob
from datetime import datetime, timedelta
import argparse
import subprocess
from string import Template
import time
import copy
import operator
import math

API_KEY = os.environ.get('SMRT_API')

if API_KEY is None:
    sys.exit('Api key not found')

smart_sheet_client = smartsheet.Smartsheet(API_KEY)
smart_sheet_client.errors_as_exceptions(True)

def create_folder(new_folder_name, location_id,location_tag):

    if location_tag == 'f':
        response = smart_sheet_client.Folders.create_folder_in_folder(str(location_id), new_folder_name)

    elif location_tag == 'w':
        response = smart_sheet_client.Workspaces.create_folder_in_workspace(str(location_id), new_folder_name)

    elif location_tag == 'h':
        response = smart_sheet_client.Home.create_folder(new_folder_name)

    return response

def create_workspace_home(workspace_name):
    # create WRKSP command
    workspace = smart_sheet_client.Workspaces.create_workspace(smartsheet.models.Workspace({'name': workspace_name}))
    return workspace

def get_sheet_list(location_id, location_tag):
    #Read in all sheets for account
    if location_tag == 'a':
        ssin = smart_sheet_client.Sheets.list_sheets(include="attachments,source,workspaces",include_all=True)
        sheets_list = ssin.data

    elif location_tag == 'f' or location_tag == 'w':
        location_object = get_object(str(location_id), location_tag)
        sheets_list = location_object.sheets

    return sheets_list

def get_folder_list(location_id, location_tag):

    if location_tag == 'f' or location_tag == 'w':
        location_object = get_object(str(location_id), location_tag)
        folders_list = location_object.folders

    elif location_tag == 'a':
        folders_list = smart_sheet_client.Home.list_folders(include_all=True)

    return folders_list

def get_workspace_list():
    # list WRKSPs command
    read_in = smart_sheet_client.Workspaces.list_workspaces(include_all=True)
    workspaces = read_in.data
    return workspaces

def get_object(object_id, object_tag):

    if object_tag == 'f':
        obj = smart_sheet_client.Folders.get_folder(str(object_id))
    elif object_tag == 'w':
        obj = smart_sheet_client.Workspaces.get_workspace(str(object_id))
    elif object_tag == 's':
        obj = smart_sheet_client.Sheets.get_sheet(str(object_id))

    return obj

def get_wo_dir_list():
    """work in progress"""

    wo_dir_list = glob.glob('[0-9][0-9][0-9][0-9][0-9][0-9][0-9]')
    return wo_dir_list

"""
MAIN: Generate issues files using jiraq and jiraq.parser
"""
mm_dd_yy = datetime.now().strftime('%m%d%y')

Wrksps = get_workspace_list()

for space in Wrksps:
    if space.name == 'QC':
        qc_space = space

for sheet in get_sheet_list(qc_space.id, 'w'):
    if 'issues.current' in sheet.name:
        current_sheet = sheet

current_sheet = get_object(current_sheet.id, 's')

ss_columns = current_sheet.columns

ss_columns_id = {}

for col in ss_columns:
    ss_columns_id[col.title] = col.id


woids = []
print('Woids (Enter "return c return" to continue): ')
while True:
    woid_in = input()

    if woid_in != 'c':
        woids.append(woid_in)
    else:
        break

with open('woids','w') as wof:
    for wo in woids:
        wof.write(wo + '\n')
print('-----------------')
print('Running jiraq...')
print('-----------------')
subprocess.run(['/bin/bash','jiraq'])

print('Running Parser...')
print('-----------------')
subprocess.run(['python3.5','jiraq_parser.py', 'woids.stats.txt'])

wo_delete = {}
for row in current_sheet.rows:
    for cel in row.cells:
        if cel.column_id == ss_columns_id['Work Order ID']:
            wo_delete[row.id] = True


with open('issue.status.{}.tsv'.format(mm_dd_yy),'r') as issues_file:
    issues_reader = csv.DictReader(issues_file, delimiter = '\t')
    header = issues_reader.fieldnames

    #Start of header check
    header_check = True
    for head in header:
        if head not in ss_columns_id.keys():
            exit('{} column not found in Smartsheet\nPlease edit the sheet reference in the QC Active Sheet'.format(head))
            new_column = smartsheet.smartsheet.models.Column({'title': head,
                                                              'type': 'TEXT_NUMBER'})
            #Add column to smartsheet
            smart_sheet_client.Sheets.add_columns(current_sheet.id,[new_column])
            ss_columns = smart_sheet_client.Sheets.get_columns(current_sheet.id, include_all = True).data


    ss_columns_id = {}
    for col in ss_columns:
        ss_columns_id[col.title] = col.id

    adding_rows = []
    updating_rows = []
    for line in issues_reader:
        wo_found = False

        #Check for WO in smartsheet
        for row in current_sheet.rows:
            for cel in row.cells:
                if cel.column_id == ss_columns_id['Work Order ID']:
                    if cel.display_value == line['Work Order ID']:
                        wo_found = True
                        wo_delete[row.id] = False
                        id = row.id



        up_ind = len(updating_rows)
        add_ind = len(adding_rows)
        #Updating or Adding?
        if wo_found:
            updating_rows.append(smart_sheet_client.models.Row())
            new_row = updating_rows[up_ind]
            new_row.id = id

        else:
            adding_rows.append(smart_sheet_client.models.Row())
            new_row = adding_rows[add_ind]
            new_row.to_bottom = True

        #update/add row with info from line
        for col in ss_columns_id:
                ind = len(new_row.cells)
                new_row.cells.append(smartsheet.models.Cell)
                new_row.cells[ind].column_id = ss_columns_id[col]
                if col in line.keys():
                    new_row.cells[ind].value = line[col]
                else:
                    new_row.cells[ind].value = 0
                new_row.cells[ind].strict = False

#Get row ids for deletion
delete_list = []
for key in wo_delete.keys():
    if wo_delete[key]:
        delete_list.append(key)

#Delete Rows
if len(delete_list) != 0:
    smart_sheet_client.Sheets.delete_rows(current_sheet.id, delete_list)

print('Updating Smartsheet...')
print('-----------------')
#Add rows
smart_sheet_client.Sheets.add_rows(current_sheet.id,adding_rows)

#Update rows
smart_sheet_client.Sheets.update_rows(current_sheet.id, updating_rows)


updated_sheet = smart_sheet_client.Sheets.update_sheet(current_sheet.id,smartsheet.models.Sheet({"name" : 'issues.current.{}'.format(mm_dd_yy)}))

print('debug statment')