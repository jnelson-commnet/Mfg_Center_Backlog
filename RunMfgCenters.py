import os
import sys
import pandas as pd

# Create system path to the FB API scripts
sys.path.insert(0, 'Z:\Python projects\FishbowlAPITestProject')
import connecttest


def run_query(dirPath, sqlFile, excelFile):
	myresults = connecttest.create_connection(dirPath, sqlFile)
	myexcel = connecttest.makeexcelsheet(myresults)
	connecttest.save_workbook(myexcel, dirPath, excelFile)


def save_sheet(backlogDF, wbName, dirPath):
	writer = pd.ExcelWriter(os.path.join(dirPath, wbName))  # Creates a test excel file
	backlogDF.to_excel(writer, sheet_name='Sheet', index=False)  # Fills the test excel with the whole timeline
	writer.save()


# Save this file's directory as a string and build other relevant paths
homey = os.path.abspath(os.path.dirname(__file__))
sqlPath = os.path.join(homey, 'SQL')
backlogFilenameTxt = 'MfgBacklog.txt'
backlogFilenameExcel = 'MfgBacklog.xlsx'



# Query FB for Mfg center info
run_query(sqlPath, backlogFilenameTxt, backlogFilenameExcel)

print('query complete ..')

# Pull query results into pandas
mfgBacklogFilepath = os.path.join(sqlPath, backlogFilenameExcel)
mfgBacklog = pd.read_excel(mfgBacklogFilepath, header=0)
### This is fixing the comma issue in tandem with some weird SQL
mfgBacklog['Customer'] = mfgBacklog['Customer'].str.replace('COMMAESCAPE', ',')
### Should probably fix the root cause of this in the API
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

print('sheets ready ..')

# Save dataFrames to Excel


rackingFilename = 'Racking_Backlog.xlsx'
proLineFilename = 'ProLine_Backlog.xlsx'
cableAssyFilename = 'Cable_Assy_Backlog.xlsx'

save_sheet(rackingBacklog, rackingFilename, homey)
save_sheet(proLineBacklog, proLineFilename, homey)
save_sheet(cableAssyBacklog, cableAssyFilename, homey)

# import email_tool

# rackingRecipientList = ['jnelson@commnetsystems.com']
# proLineRecipientList = ['jnelson@commnetsystems.com']

# email_tool.send_email(rackingRecipientList, rackingFilename)
# email_tool.send_email(proLineRecipientList, proLineFilename)
# email_tool.send_email(cableAssyRecipientList, cableAssyFilename)

print('done!')

