# Auto_Issues
Automatiaclly update the issues.current and QC Active Issues sheets in smartsheet

Protocol for use:

First download csv:
https://jira.ris.wustl.edu/issues/?filter=13180#
on the terminal, log in to server,


(Put gsub_alt somewhere accessible)
gsub_alt

cd /gscmnt/gc2783/qc/GMSworkorders

cat > woids
(highlight work order ID column and copy and paste)
control D to save the file

python3.5 ~awollam/aw/auto_issues.py

(output in terminal: summary of zero builds or all succeeded builds)
output file: issue....tsv

File is automatically uploaded to smartsheet
