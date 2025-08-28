import sys
from datetime import datetime

#FOR EXCEL VERSION YOU NEED INSTALL THIS LIBRARY. "pip install XlsxWriter"
import xlsxwriter

#FOR MULTIPLE ARGUMENTS "pip install argparse"
import argparse

#SET THROWLOG ON(1) / OFF(0)
enablethrowlog = '0'

#FOR DEBUG ONLY
#sys.argv = ['./logfilter_2.py', 'in.trc']

parser = argparse.ArgumentParser()
parser.add_argument('file', type=argparse.FileType('rt'), nargs='*')

#Creating the xls file.
workbook = xlsxwriter.Workbook(sys.argv[1] +'.xlsx')

#Creating the sheets
xls_ahlog = workbook.add_worksheet('AuctionHouse_Log')
xls_pslog = workbook.add_worksheet('PersonalShop_Log')
xls_tradelog = workbook.add_worksheet('Trade_Log')
xls_gwhlog = workbook.add_worksheet('GuildWarehouse_Log')
xls_maillog = workbook.add_worksheet('Mail_Log')
xls_throwlog = workbook.add_worksheet('Throw_Log')
xls_entrylog = workbook.add_worksheet('No_Entry_Hack_Log')

#Creating the sheet headers and line counters.   
#PS_Log Headers
xls_pslog.freeze_panes(1, 0)
xls_pslog.write_string(0, 0, 'SellerCharIdx')
xls_pslog.write_string(0, 1, 'BuyerCharIDX')
xls_pslog.write_string(0, 2, 'ItemKind')
xls_pslog.write_string(0, 3, 'ItemOpt')
xls_pslog.set_column(3, 3, 12)
xls_pslog.write_string(0, 4, 'AlzPrice')
xls_pslog.set_column(4, 4, 15)
xls_pslog_counter = 1

#Trade_Log Headers
xls_tradelog.freeze_panes(1, 0)
xls_tradelog.write_string(0, 0, 'TimeStamp')
xls_tradelog.set_column(0, 0, 12)
xls_tradelog.write_string(0, 1, 'SrcCharIDX')
xls_tradelog.write_string(0, 2, 'DesCharIDX')
xls_tradelog.write_string(0, 3, 'ItemKind')
xls_tradelog.write_string(0, 4, 'ItemOpt')
xls_tradelog.set_column(4, 4, 12)
xls_tradelog.write_string(0, 5, 'Alz')
xls_tradelog.set_column(5, 5, 15)
xls_tradelog_counter = 1

#AH_Log Headers
xls_ahlog.freeze_panes(1, 0)
xls_ahlog.write_string(0, 0, 'BuyerCharIdx')
xls_ahlog.write_string(0, 1, 'SellerCharIdx')
xls_ahlog.write_string(0, 2, 'ItemKind')
xls_ahlog.write_string(0, 3, 'ItemOpt')
xls_ahlog.set_column(3, 3, 12)
xls_ahlog.write_string(0, 4, 'AlzPriceEach')
xls_ahlog.set_column(4, 4, 15)
xls_ahlog.write_string(0, 5, 'Count')
xls_ahlog.write_string(0, 6, 'TotalPrice')
xls_ahlog.set_column(6, 6, 15)
xls_ahlog_counter = 1

#GWH_Log Headers
xls_gwhlog.freeze_panes(1, 0)
xls_gwhlog.write_string(0, 0, 'GuildNo')
xls_gwhlog.write_string(0, 1, 'CharIDX')
xls_gwhlog.write_string(0, 2, 'In/Out')
xls_gwhlog.write_string(0, 3, 'ItemKind')
xls_gwhlog.write_string(0, 4, 'ItemOpt')
xls_gwhlog.set_column(4, 4, 12)
xls_gwhlog.write_string(0, 5, 'Count')
xls_gwhlog.write_string(0, 6, 'AlzAmount')
xls_gwhlog.set_column(6, 6, 12)
xls_gwhlog_counter = 1

#Mail_Log Headers
xls_maillog.freeze_panes(1, 0)
xls_maillog.write_string(0, 0, 'TimeStamp')
xls_maillog.set_column(0, 0, 12)
xls_maillog.write_string(0, 1, 'FromCharIDX')
xls_maillog.write_string(0, 2, 'ToCharIDX')
xls_maillog.write_string(0, 3, 'ItemKind')
xls_maillog.write_string(0, 4, 'ItemOpt')
xls_maillog.set_column(4, 4, 12)
xls_maillog.write_string(0, 5, 'AlzAmount')
xls_maillog.set_column(5, 5, 15)
xls_maillog.write_string(0, 6, 'ReceivedMailID')
xls_maillog.set_column(6, 6, 12)
xls_maillog_counter = 1

#Throw_Log Headers
xls_throwlog.freeze_panes(1, 0)
xls_throwlog.write_string(0, 0, 'CharacterIDX')
xls_throwlog.write_string(0, 1, 'ItemKind')
xls_throwlog.write_string(0, 2, 'ItemOpt')
xls_throwlog.set_column(2, 2, 15)
xls_throwlog.write_string(0, 3, 'Throw/Pickup')
xls_throwlog_counter = 1

#No_Entry_Hack_Log Headers
xls_entrylog.freeze_panes(1, 0)
xls_entrylog.write_string(0, 0, 'TimeStamp')
xls_entrylog.write_string(0, 1, 'CharacterIdx')
xls_entrylog.write_string(0, 2, 'Action')
xls_entrylog_counter = 1

args = parser.parse_args()
dataLog = []
for f in args.file:
    for line in f:
        line = line.replace("\n","")
        splittedline = line.split("|")

        #Personal Shop Logger
        if splittedline[1] == '5115':
            xls_pslog.write_string(xls_pslog_counter, 0, splittedline[2])
            xls_pslog.write_string(xls_pslog_counter, 1, splittedline[6])
            xls_pslog.write_string(xls_pslog_counter, 2, splittedline[3])
            xls_pslog.write_string(xls_pslog_counter, 3, splittedline[4])
            xls_pslog.write_string(xls_pslog_counter, 4, splittedline[7])
            xls_pslog_counter += 1

        #Trade Item Logger
        if splittedline[1] == '5131':
            xls_tradelog.write_string(xls_tradelog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_tradelog.write_string(xls_tradelog_counter, 1, splittedline[2])
            xls_tradelog.write_string(xls_tradelog_counter, 2, splittedline[6])
            xls_tradelog.write_string(xls_tradelog_counter, 3, splittedline[3])
            xls_tradelog.write_string(xls_tradelog_counter, 4, splittedline[4])
            xls_tradelog.write_string(xls_tradelog_counter, 5, '-')
            xls_tradelog_counter += 1

        #Trade Alz Logger
        if splittedline[1] == '5203':
            xls_tradelog.write_string(xls_tradelog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_tradelog.write_string(xls_tradelog_counter, 1, splittedline[2])
            xls_tradelog.write_string(xls_tradelog_counter, 2, splittedline[5])
            xls_tradelog.write_string(xls_tradelog_counter, 3, '-')
            xls_tradelog.write_string(xls_tradelog_counter, 4, '-')
            xls_tradelog.write_string(xls_tradelog_counter, 5, splittedline[3])
            xls_tradelog_counter += 1

        #Auction House Logger
        if splittedline[1] == '51044':
            totalprice = str(int(splittedline[8]) * int(splittedline[9]))
            xls_ahlog.write_string(xls_ahlog_counter, 0, splittedline[2])
            xls_ahlog.write_string(xls_ahlog_counter, 1, splittedline[7])
            xls_ahlog.write_string(xls_ahlog_counter, 2, splittedline[3])
            xls_ahlog.write_string(xls_ahlog_counter, 3, splittedline[4])
            xls_ahlog.write_string(xls_ahlog_counter, 4, splittedline[8])
            xls_ahlog.write_string(xls_ahlog_counter, 5, splittedline[9])
            xls_ahlog.write_string(xls_ahlog_counter, 6, totalprice)
            xls_ahlog_counter += 1
        
        #Guild Warehouse Item Logger
        if splittedline[1] == '51049':
            xls_gwhlog.write_string(xls_gwhlog_counter, 0, splittedline[6])
            xls_gwhlog.write_string(xls_gwhlog_counter, 1, splittedline[2])
            if splittedline[7] == '0':
                xls_gwhlog.write_string(xls_gwhlog_counter, 2, 'In')
            else:
                xls_gwhlog.write_string(xls_gwhlog_counter, 2, 'Out')
            xls_gwhlog.write_string(xls_gwhlog_counter, 3, splittedline[3])
            xls_gwhlog.write_string(xls_gwhlog_counter, 4, splittedline[4])
            xls_gwhlog.write_string(xls_gwhlog_counter, 5, splittedline[9])
            xls_gwhlog.write_string(xls_gwhlog_counter, 6, '-')
            xls_gwhlog_counter += 1

        #Guild Warehouse Alz Logger
        if splittedline[1] == '10953':
            xls_gwhlog.write_string(xls_gwhlog_counter, 0, splittedline[3])
            xls_gwhlog.write_string(xls_gwhlog_counter, 1, splittedline[2])
            if splittedline[4] == '0':
                xls_gwhlog.write_string(xls_gwhlog_counter, 2, 'In')
            else:
                xls_gwhlog.write_string(xls_gwhlog_counter, 2, 'Out')
            xls_gwhlog.write_string(xls_gwhlog_counter, 3, '-')
            xls_gwhlog.write_string(xls_gwhlog_counter, 4, '-')
            xls_gwhlog.write_string(xls_gwhlog_counter, 5, '-')
            xls_gwhlog.write_string(xls_gwhlog_counter, 6, splittedline[6])
            xls_gwhlog_counter += 1

        #Mail Item Logger
        if splittedline[1] == '51019':
            xls_maillog.write_string(xls_maillog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_maillog.write_string(xls_maillog_counter, 1, splittedline[2])
            xls_maillog.write_string(xls_maillog_counter, 2, splittedline[8])
            xls_maillog.write_string(xls_maillog_counter, 3, splittedline[3])
            xls_maillog.write_string(xls_maillog_counter, 4, splittedline[4])
            xls_maillog.write_string(xls_maillog_counter, 5, '-')
            xls_maillog.write_string(xls_maillog_counter, 6, splittedline[9])
            xls_maillog_counter += 1

        #Mail Alz Logger
        if splittedline[1] == '5361':
            xls_maillog.write_string(xls_maillog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_maillog.write_string(xls_maillog_counter, 1, splittedline[2])
            xls_maillog.write_string(xls_maillog_counter, 2, splittedline[6])
            xls_maillog.write_string(xls_maillog_counter, 3, '-')
            xls_maillog.write_string(xls_maillog_counter, 4, '-')
            xls_maillog.write_string(xls_maillog_counter, 5, splittedline[3])
            xls_maillog.write_string(xls_maillog_counter, 6, splittedline[7])
            xls_maillog_counter += 1

        #Throw/Pickup Logger
        #THIS TAKE TONS OF TIME To RUN, ENABLE IT IF REALLY NEEDED!
        if enablethrowlog == '1':
            if splittedline[1] == '5101' or splittedline[1] == '5102':
                xls_throwlog.write_string(xls_throwlog_counter, 0, splittedline[2])
                xls_throwlog.write_string(xls_throwlog_counter, 1, splittedline[3])
                xls_throwlog.write_string(xls_throwlog_counter, 2, splittedline[4])
                if splittedline[1] == '5101':
                    xls_throwlog.write_string(xls_throwlog_counter, 3, 'Throw')
                elif splittedline[1] == '5102':
                    xls_throwlog.write_string(xls_throwlog_counter, 3, 'Pickup')
                xls_throwlog_counter += 1

        #NoEntryHackDetector
        # if splittedline[1] == '5104':
        #     xls_entrylog.write_string(xls_entrylog_counter, 0, splittedline[0])
        #     xls_entrylog.write_string(xls_entrylog_counter, 1, splittedline[2])
        #     action = "Move Item Id: " + splittedline[3] + " ItemOpt: " + splittedline[4] + " From: " + splittedline[6] + " slot."
        #     xls_entrylog.write_string(xls_entrylog_counter, 2, action)
        #     xls_entrylog_counter += 1
        #
        # if splittedline[1] == '5105':
        #     xls_entrylog.write_string(xls_entrylog_counter, 0, splittedline[0])
        #     xls_entrylog.write_string(xls_entrylog_counter, 1, splittedline[2])
        #     action = "Move Item Id: " + splittedline[3] + " ItemOpt: " + splittedline[4] + " To: " + splittedline[6] + " slot."
        #     xls_entrylog.write_string(xls_entrylog_counter, 2, action)
        #     xls_entrylog_counter += 1

        if splittedline[1] == '51022':
            xls_entrylog.write_string(xls_entrylog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_entrylog.write_string(xls_entrylog_counter, 1, splittedline[2])
            action = "Dungeon entry used: " + splittedline[3] + "-" + splittedline[4] + ". Slot: " + splittedline[7] + " Dungeon: " + splittedline[8] + "."
            xls_entrylog.write_string(xls_entrylog_counter, 2, action)
            xls_entrylog_counter += 1

        if splittedline[1] == '6167':
            xls_entrylog.write_string(xls_entrylog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_entrylog.write_string(xls_entrylog_counter, 1, splittedline[2])
            action = "Dungeon: " + splittedline[3] + " started."
            xls_entrylog.write_string(xls_entrylog_counter, 2, action)
            xls_entrylog_counter += 1

        if splittedline[1] == '9':
            xls_entrylog.write_string(xls_entrylog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_entrylog.write_string(xls_entrylog_counter, 1, "-")
            action = "Disconnect from IP: " + splittedline[2] + "."
            xls_entrylog.write_string(xls_entrylog_counter, 2, action)
            xls_entrylog_counter += 1

        if splittedline[1] == '9103':
            xls_entrylog.write_string(xls_entrylog_counter, 0, datetime.fromtimestamp(int(splittedline[0])).strftime('%Y-%m-%d %H:%M:%S'))
            xls_entrylog.write_string(xls_entrylog_counter, 1, "-")
            action = "Characteridx: " + splittedline[2] + " entered the channel."
            xls_entrylog.write_string(xls_entrylog_counter, 2, action)
            xls_entrylog_counter += 1

#Closing the Excel file.
workbook.close()