__author__ = "Thomas Antonacci"


"""
This script will now be designed to update and add work orders to/in smartsheet from jira csv/tsv

This script relies on the jiraq bash script and ~awollam/aw/jiraq_parser.py woids.stats.txt
Meant to be run on personal gsub_alt(uses registry.gsc.wustl.edu/antonacci.t.j/genome_perl_environment:latest)

TODO: Notation and clean up
"""


import smartsheet
import csv
import os
import sys
from datetime import datetime
import subprocess



API_KEY = os.environ.get('SMRT_API')

if API_KEY is None:
    sys.exit('Api key not found')

smart_sheet_client = smartsheet.Smartsheet(API_KEY)
smart_sheet_client.errors_as_exceptions(True)

run_from = os.getcwd()
os.chdir('/gscmnt/gc2783/qc/GMSworkorders')

"""
Smartsheet tools
-------------------------
"""
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

"""
Get sheets and columns needed from smartsheet
"""
mm_dd_yy = datetime.now().strftime('%m%d%y')

Wrksps = get_workspace_list()

#Get qc workspace
for space in Wrksps:
    if space.name == 'QC':
        qc_space = space

#get current issues sheet
for sheet in get_sheet_list(qc_space.id, 'w'):
    if 'issues.current' in sheet.name:
        current_sheet = sheet

current_sheet = get_object(current_sheet.id, 's')

#Get QC Active sheet
for sheet in get_sheet_list(qc_space.id, 'w'):
    if 'QC Active Issues' == sheet.name:
        qc_active_sheet = sheet

qc_active_sheet = get_object(qc_active_sheet.id, 's')

#Get QC Active columns
active_columns = qc_active_sheet.columns
active_columns_id = {}

for col in active_columns:
    active_columns_id[col.title] = col.id

#Get current sheet columns
ss_columns = current_sheet.columns

ss_columns_id = {}

for col in ss_columns:
    ss_columns_id[col.title] = col.id
    
#Input sheet from Jira

jira_sheet = []
print('Paste Jira Sheet ("return c return" to continue): ')

#Input woids from Jira

woids = []
print('Woids (Enter "return c return" to continue): ')
while True:
    woid_in = input()

    if woid_in != 'c':
        woids.append(woid_in)
    else:
        break

#check ss woids and jira woids
active_wos = []
for row in qc_active_sheet.rows[7:]:
    resolved = False
    for cell in row.cells:
        if cell.column_id == active_columns_id['Health'] and cell.value == 'Blue':
            resolved = True

    for cell in row.cells:
        if cell.column_id == active_columns_id['Work Order ID'] and not resolved:
            active_wos.append(cell.value)
        elif str(cell.value).replace('.0','') in woids and resolved:
            print('{} found resolved in QC Active Issues.'.format(str(cell.value).replace('.0','')))

for woid in active_wos:
    active_wos[active_wos.index(woid)] = str(woid).replace('.0','')

for woid in woids:
    if woid not in active_wos:
        print('{} found in Jira but not Smartsheet.'.format(woid))



jira_temp = 'jira_temp.tsv'
with open(jira_temp, 'w') as js:
    while True:
        sheet_in = input()

        if sheet_in != 'c':
            js.write(sheet_in + '\n')
        else:
            break

#eliminate dup linked issues column
with open(jira_temp, 'r') as jt, open(jira_temp + '_1', 'w') as jt1:
    jira_read = csv.reader(jt, delimiter = '\t')
    temp_writer = csv.writer(jt1, delimiter = '\t')
    data = [r for r in jira_read]
    
    dup_found = False
    i = 0
    j = 0
    for title in data[0]:
        if title == 'Outward issue link (Depends)' and not dup_found:
            dup_found = True
        elif title == 'Outward issue link (Depends)' and dup_found:
            data[0][i] = title + '_{}'.format(j)
            j +=1
        i += 1
        
    for rw in data:
        temp_writer.writerow(rw)
      
os.rename(jira_temp + '_1', jira_temp)


#Update Smartsheet QC Active Issues with new work orders

row_num = len(qc_active_sheet.rows) + 1
woids = []
resolved_woids = []

with open(jira_temp, 'r') as jt:
    jira_read = csv.DictReader(jt, delimiter = '\t')
    header = jira_read.fieldnames
    for line in jira_read:
        try:
            woids.append(line['Custom field (Work Order ID)'])
        except ValueError:
            print('No Work Order ID for {}.'.line['Issue key'])
    jt.seek(1)


    #check ss woids and jira woids
    active_wos = []
    for row in qc_active_sheet.rows[7:]:
        resolved = False
        for cell in row.cells:
            if cell.column_id == active_columns_id['Health'] and cell.value == 'Blue':
                resolved = True

        for cell in row.cells:
            if cell.column_id == active_columns_id['Work Order ID'] and not resolved:
                active_wos.append(cell.value)
            elif str(cell.value).replace('.0','') in woids and resolved:
                print('{} found resolved in QC Active Issues.'.format(str(cell.value).replace('.0','')))
                resolved_woids.append(resolved_woids.append(str(cell.value).replace('.0','')))

    for woid in active_wos:
        active_wos[active_wos.index(woid)] = str(woid).replace('.0','')

    for woid in woids:
        if (woid not in active_wos) and (woid not in resolved_woids):
            print('{} found in Jira but not Smartsheet.'.format(woid))
            print('Adding row to smartsheet...')
            for line in jira_read:
                if line['Custom field (Work Order ID)'] == woid:

                    new_row = smartsheet.smartsheet.models.Row()

                    #Date run
                    std_yyyymmdd = datetime.now().strftime('%Y-%m-%d')
                    new_row.cells.append({'column_id': active_columns_id['Auto Issues Last Updated'], 'value': std_yyyymmdd})
      
                    #Jira information
                    new_row.cells.append({'column_id': active_columns_id['Work Order ID'], 'value': int(line['Custom field (Work Order ID)'])})
                    new_row.cells.append({'column_id': active_columns_id['Component/s'],'value': line['Component/s']})
                    new_row.cells.append({'column_id': active_columns_id['Labels'], 'value': line['Labels']})
                    new_row.cells.append({'column_id': active_columns_id['Summary'], 'value': line['Summary']})
                    new_row.cells.append({'column_id': active_columns_id['Issue Key'], 'value': line['Issue key']})

                    #formulas
                    new_row.cells.append({'column_id': active_columns_id['Health'], 'formula': '=IFERROR(IF([QC Complete?]{num} = 1, "Blue", IF([Total Builds]{num} = [Succeeded Builds]{num}, IF([Total Builds]{num} = 0, "Red", "Green"), IF([Build Failed]{num} > 0, "Red", "Yellow"))), "Red")'.format(num=row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Failed Flag'], 'formula': '=IFERROR(IF([Build Failed]{} > 0, 1, 0), 0)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Total Builds'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 2, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Succeeded Builds'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 3, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Scheduled'],'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 4, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Running Builds'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 5, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Build Needed'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 6, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Build Failed'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 7, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Build Requested'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 8, false)'.format(row_num)})
                    new_row.cells.append({'column_id': active_columns_id['Unstartable Builds'], 'formula': '=VLOOKUP($[Work Order ID]{}, {{issue.status.060319-2 Range 2}}, 9, false)'.format(row_num)})


                    #Hyperlinks
                    conf_url = 'https://confluence.ris.wustl.edu/pages/viewpage.action?spaceKey=AD&title=WorkOrder+{}'.format(str(line['Custom field (Work Order ID)']).replace('.0',''))
                    jira_url = 'https://jira.ris.wustl.edu/browse/{}'.format(line['Issue key'])

                    new_row.cells.append({'column_id': active_columns_id['Confluence Page WOID'], 'value': conf_url, 'hyperlink': {'url' : conf_url}})
                    new_row.cells.append({'column_id': active_columns_id['JIRA Issue Link'], 'value': jira_url, 'hyperlink': {'url': jira_url}})


                    #Blanks
                    new_row.cells.append({'column_id': active_columns_id['QC Queried Date'], 'value': ''})
                    new_row.cells.append({'column_id': active_columns_id['QC Complete?'], 'value': False})
                    new_row.cells.append({'column_id': active_columns_id['internal comment about this issue'], 'value': ''})
                    new_row.cells.append({'column_id': active_columns_id['Analysis Project Status'], 'value': ''})
                    new_row.cells.append({'column_id': active_columns_id['Weekly Update (PM)'], 'value': ''})
                    new_row.cells.append({'column_id': active_columns_id['Analyst'], 'value': ''})



                    #Linked Issues
                    linked_issues = []
                    for key in line.keys():
                        if 'Outward issue link (Depends)' in key:
                            linked_issues.append(line[key])

                    new_row.cells.append({'column_id': active_columns_id['Linked JIRA Parent/Dependent Issues'], 'value': ','.join(linked_issues)})
                    #add row?

                    smart_sheet_client.Sheets.add_rows(qc_active_sheet.id, [new_row])

            row_num += 1
            jt.seek(1)

        elif woid in active_wos:
            #update linked jira issue info
            up_row = smartsheet.smartsheet.models.Row()

            for row in qc_active_sheet.rows[7:]:
                if str(row.get_column(active_columns_id['Work Order ID']).value).replace('.0','') == woid:
                    up_row.id = row.id

            for line in jira_read:
                if line['Custom field (Work Order ID)'] == woid:
                    linked_issues = []
                    for key in line.keys():
                        if 'Outward issue link (Depends)' in key and line[key] != '':
                            linked_issues.append(line[key])

                    std_yyyymmdd = datetime.now().strftime('%Y-%m-%d')

                    up_row.cells.append({'column_id': active_columns_id['Linked JIRA Parent/Dependent Issues'], 'value': ','.join(linked_issues)})
                    up_row.cells.append({'column_id': active_columns_id['Auto Issues Last Updated'], 'value': std_yyyymmdd})

                    smart_sheet_client.Sheets.update_rows(qc_active_sheet.id, [up_row])
            jt.seek(1)

#write woids file for jiraq
with open('woids','w') as wof:
    for wo in woids:
        wof.write(wo + '\n')


print('-----------------')
print('Running jiraq...')
subprocess.run(['/bin/bash','jiraq'])
print('-----------------')


print('Running Parser...')
subprocess.run(['python3.5','jiraq_parser.py', 'woids.stats.txt'])

#Get work orders to delete
wo_delete = {}
for row in current_sheet.rows:
    for cel in row.cells:
        if cel.column_id == ss_columns_id['Work Order ID']:
            wo_delete[row.id] = True

#Update current issues file with new info from jiraq
with open('issue.status.{}.tsv'.format(mm_dd_yy),'r') as issues_file:
    issues_reader = csv.DictReader(issues_file, delimiter = '\t')
    header = issues_reader.fieldnames

    #Start of header check
    header_check = True
    for head in header:
        if head not in ss_columns_id.keys():
            print('{} column not found in Smartsheet\nPlease check the sheet reference in the QC Active Sheet'.format(head))
            new_column = smartsheet.smartsheet.models.Column({'title': head,
                                                              'type': 'TEXT_NUMBER',
                                                              'index': len(ss_columns_id) - 1})
            #Add column to smartsheet
            smart_sheet_client.Sheets.add_columns(current_sheet.id,[new_column])
            ss_columns = smart_sheet_client.Sheets.get_columns(current_sheet.id, include_all = True).data

            new_column = smartsheet.smartsheet.models.Column({'title': head,
                                                              'type': 'TEXT_NUMBER',
                                                              'index': len(qc_active_sheet.columns) - 1})

            col_resp = smart_sheet_client.Sheets.add_columns(qc_active_sheet.id, [new_column]).data



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
            #Must be added above the last row(below the 2nd last) in order for the cross sheet reference to work
            new_row.sibling_id = current_sheet.rows[-2].id

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
print('-----------------')

print('Updating Smartsheet...')
print('-----------------')
#Add rows
smart_sheet_client.Sheets.add_rows(current_sheet.id,adding_rows)

#Update rows
smart_sheet_client.Sheets.update_rows(current_sheet.id, updating_rows)

#Update current issues sheet with today's date
updated_sheet = smart_sheet_client.Sheets.update_sheet(current_sheet.id,smartsheet.models.Sheet({"name" : 'issues.current.{}'.format(mm_dd_yy)}))

#Delete temp file and move used issues file to issues.archive
os.remove(jira_temp)

os.rename('issue.status.{}.tsv'.format(mm_dd_yy), 'issues.archive/issue.status.{}.tsv'.format(mm_dd_yy))
os.chdir(run_from)
