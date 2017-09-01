import os
import sys
import pandas as pd

# Create system path to the FB API scripts
sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')
import connecttest

# Save this file's directory as a string and build other relevant paths
homey = os.path.abspath(os.path.dirname(__file__))
sqlPath = os.path.join(homey, 'SQL')
backlogFilename = 'MfgBacklog.txt'

def run_queries():
	myresults = connecttest.create_connection(sqlPath, backlogFilename)
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, sqlPath, backlogFilename)



# Query FB for Mfg center info
run_queries()

# Pull query results into pandas
mfgBacklogFilepath = os.path.join(sqlPath, backlogFilename)
mfgBacklog = pd.read_excel(mfgBacklogFilepath, header=0)
mfgBacklog.sort_values(by='PartNum', ascending=True, inplace=True)

# Split dataFrames by Mfg Center
rackingBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Racking'].copy()
proLineBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Pro line'].copy()
# pcbBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'PCB'].copy()
# shippingBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Shipping'].copy()
# cableAssyBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Cable Assembly'].copy()
# kittingBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Kitting'].copy()
# plasticsBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Plastics'].copy()
# labelsBacklog = mfgBacklog[mfgBacklog['Mfg Center'] == 'Labels'].copy()

# Save dataFrames to Excel

def save_sheets(backlogDF, wbName):
	writer = pd.ExcelWriter(os.path.join(homey, wbName))  # Creates a test excel file
	backlogDF.to_excel(writer, 'Sheet')  # Fills the test excel with the whole timeline
	writer.save()

rackingFilename = 'Racking_Backlog.xlsx'
proLineFilename = 'ProLine_Backlog.xlsx'

save_sheets(rackingBacklog, rackingFilename)
save_sheets(proLineBacklog, proLineFilename)

# import email_tool

# rackingRecipientList = ['jnelson@commnetsystems.com']
# proLineRecipientList = ['jnelson@commnetsystems.com']

# email_tool.send_email(rackingRecipientList, rackingFilename)
# email_tool.send_email(proLineRecipientList, proLineFilename)

