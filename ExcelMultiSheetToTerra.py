#!env python3
#Parse Excel file with multiple sheets and convert to terraform format
#Author : Manjesh.munegowda@sap.com
#Date : Oct-04-2019
#Oct-09-2019: as per agreed upon format, need to change sheet names and add Udr & Vnet-peer
#Oct-15-2019: Converted script to python3.6 version, added Global-Var, new input file:network-tffile_v1.1.xlsx, modified Headers to the new input file
#             Implemented terraform.tfvars exists and rename with date & timestamp
#Oct-16-2019: Fixed Vnet address space issue, added inputFile exist and write permission check, Fixed single blank space in input file 
#Oct-17-2019: Added write to ouputfile, script runtime time.
#Oct-18-2019: With network-tffile_v1.5.xlsx, number of columns have increased to more than 4 column in vnet address space, need to handle dynamic number of columns
#Oct-21-2019: fixed the dynamic column, created class for Brackets, comma, doubleQuote
#Oct-21-2019: Replaced static brackets with class brackets
#
#Oct-31-2019: New requirment, ask for input file and if previous terraform output file exist, show the diff after genearting the new terrafrom output file
#Nov-07-2019: Implemented - ask for input file 
#Nov-08-2019: Implemented - if previous terraform output file exist, show the diff after genearting the new terrafrom output file
#
#Mar-09-2020: New Requirement: add azurerm_network_security_rule" = (E through N on NSG tab)
#
import argparse                                                                      #Required for handling command line arguments
import pandas as pd                                                                  #Required for processing Excel file
import time                                                                          #Required for getting time 
import pathlib                                                                       #Required for finding path
import os                                                                            #Required for finding the basename 
import sys                                                                           #Required for redirecting stdout to file
import difflib                                                                       #Required for checking the difference between two files
import numpy as nump                                                                 #As of 2022 pandas deprecated numpy
startTime=time.time()
global bkp
bkp=''
#####
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
# class for brackets, instead of static brackets, comma, double quote and equal
#
class myBracket:
    def openSquare():
        return "["
    def closeSquare():
        return "]"
    def openCurl():
        return "{"
    def closeCurl():
        return "}"
    def equal():
        return "="
    def dQuote():
        return '"'
    def comma():
        return ','
#----------
osq=myBracket.openSquare()
csq=myBracket.closeSquare()
ocrl=myBracket.openCurl()
ccrl=myBracket.closeCurl()
eql=myBracket.equal()
dqt=myBracket.dQuote()
cma=myBracket.comma()
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
# Open the standardized Input file : network-tffile_v1.5.xlsx and create DataFrame for the worksheets in the Excel file
#

parser = argparse.ArgumentParser(description='Excel input parser to generate terraform.tfvars.')  #get the arguments passed at command line
parser.add_argument('inputFile', type=argparse.FileType('r'), 
                    help='Example: network-tffile_v1.5.xlsx')
args=parser.parse_args()

#print(args.inputFile.name)
#print(str(args['name']))
try:
      scriptDirectory = str(os.path.dirname(os.path.realpath(__file__)))             #Get the path to script
      scriptName=str(os.path.basename(__file__))                                     #Get the script Name
      with open(args.inputFile.name) as iFH:                                         #try to open the file passed as argument, along with -d argument
           inputFile=pathlib.Path(args.inputFile.name)
           xls=pd.ExcelFile(inputFile)
           df0=pd.read_excel(xls,'Global-Var')
           df1=pd.read_excel(xls,'Rg-Name')
           df2=pd.read_excel(xls,'Vnet-Maping')
           df3=pd.read_excel(xls,'Vnet-Add-Space')
           df4=pd.read_excel(xls,'Sub-Vnet-Map')
           df5=pd.read_excel(xls,'Pub-IP')
           df6=pd.read_excel(xls,'Nsg')
           df7=pd.read_excel(xls,'Udr')
           df8=pd.read_excel(xls,'Vnet-Peer')
except:
     error=sys.exc_info()[1]
     #print("\n\nFatal Error: {} : not found in current directory [{}], Please make sure it exists in [{}] directory and try again...\n\n".format(str(myArgs['-d']), os.getcwd(),scriptDirectory))
     print(error)
     exit()
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
# Check if terraform.tfvars exist, if exists rename with date & timestamp, and open output file for writing
#
fp=pathlib.Path('terraform.tfvars')
if fp.exists():
   if (os.access(scriptDirectory, os.W_OK)):
      bkp=time.strftime("%b-%d-%Y-%H:%M:%S%p")
      bkp="terraform.tfvars-" + bkp
      os.rename(fp, bkp)                                                        #Renaming the existing terraform.tfvars file 
      #print("Previous File : " + str(bkp))
      fp=open('terraform.tfvars','w')
   else:
      print("\n\nFatal Error: you don't have write access to write {} into {} \n\n".format(fp,scriptDirectory))
      quit()
else:
   if (os.access(scriptDirectory, os.W_OK)):
      fp=open('terraform.tfvars','w')
   else:
      print("\n\nFatal Error: you don't have write access to write {} into {} \n\n".format(fp,scriptDirectory))
      quit()
  
print("** Generating terraform.tfvars")
ctime=time.strftime("%b-%d-%Y-%H:%M:%S%p")
#sys.stdout=fp
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
# Global-Var
#
df0=df0.fillna('')
#df0=df0.strip()
for h, (index, row) in enumerate(df0.iterrows()):
    if str(row['DC'])=="":
       continue
    else:
       print("#Terraform.tfvars generated on {} using {}".format(ctime,scriptName))
       print("#Global Variables")
       print("dc\t\t{} \t{}{}{}".format(eql,dqt,str(int(row['DC'])),dqt))
       print("location\t{} \t{}{}{}".format(eql,dqt,row['Location'].strip(),dqt))
       print("tags\t\t{} \t{}{}{}".format(eql,dqt,row['Tags'].strip(),dqt))

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
# Resource Group
#
#below to getrid of NaN
#df1['resource group'].replace(' ', pd.np.nan, inplace=True)
df1['resource group'].replace(' ', nump.nan, inplace=True)
df1.dropna(subset=['resource group'], inplace=True)
df1=df1.fillna('')
print("\n\n#Input to create the ResouceGroups" )
print("\n{}resource_groups{} {} \n{}".format(dqt,dqt,eql,osq))

for i, (index, row) in enumerate(df1.iterrows()):
    if i == len(df1) - 1:
       if (row['resource group'].strip()) == "": 
          continue
       else:
          print("\t{} rg_name{}{}{}{} {}".format(ocrl,eql,dqt,row['resource group'].strip(),dqt,ccrl))
    else:
       if (row['resource group'].strip()) == "": 
          continue
       else:
          print("\t{} rg_name{}{}{}{} {}{}".format(ocrl,eql,dqt,row['resource group'].strip(),dqt,ccrl,cma))
print("{}".format(csq))

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
# Vnet Mapping
#
#below to getrid of NaN
#df2['resource group'].replace(' ', pd.np.nan, inplace=True)
df2['resource group'].replace(' ', nump.nan, inplace=True)
df2.dropna(subset=['resource group'], inplace=True)
df2=df2.fillna('')
print("\n\n#Input to create the VirtualNetworks" )
print("\n\n{}vnet_mapping{} {}\n{}".format(dqt,dqt,eql,osq))
for j, (index, row) in enumerate(df2.iterrows()):
    if j == len(df2) - 1:
       print("\t{} rg_name{}{}{}{} vnet_name{}{}{}{} {}".format(ocrl,eql,dqt,row['resource group'].strip(),dqt,eql,dqt,row['vnet name'].strip(),dqt,ccrl))
    else:
       print("\t{} rg_name{}{}{}{} vnet_name{}{}{}{} {}{}".format(ocrl,eql,dqt,row['resource group'].strip(),dqt,eql,dqt,row['vnet name'].strip(),dqt,ccrl,cma))
print("{}".format(csq))

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#vnet address space
#
print("\n\n#Input to create the subnets" )
print("\n\n{}vnet_address_space{} {}\n{}".format(dqt,dqt,eql,ocrl))

subnetCols= [ col for col in df3.columns if 'subnet' in col ]                      #Build a list of the header, for subnet

#Below filter is to get the count of variable number of subnet columns, this will add a count field to end of the DataFrame, count is used to check the last column, in below while loop
df3['count'] = df3.filter(regex="^subnet").count(axis=1)
df3.applymap(str)                                                                 #make everything a string
#df3['vnet name'].replace(' ', pd.np.nan, inplace=True)                            #Replace whitespace to NaN, in Vnet name column
df3['vnet name'].replace(' ', nump.nan, inplace=True)                            #Replace whitespace to NaN, in Vnet name column
df3.replace(' ', nump.nan, inplace=True)                                         #Replace whitespace to NaN, in all data 
df3.dropna(subset=['vnet name'], inplace=True)                                    #If Vnet Name has Nan, drop the row
#print (df3)

for k, (index, row) in enumerate(df3.iterrows()):
     if (row['vnet name'].strip()) == "":
        continue
     else:
       print("\t{}{}{} ".format(row['vnet name'].strip(),eql,osq), end="" )
       subCount=0
       colCount=int(row['count']) 						  #for each row get the count of subnet columns, since subnet column is dynamic 
       #print("ColCount : ", str(colCount))
       while subCount <= colCount:                                                #Loop thru the subnet count 
          for clmn in subnetCols:                                                 #Loop thru the subnet headers
            if (subCount == colCount-1) or subnetCols[-1] == clmn:	          #Check if last column, if yes - don't add the comma to subnet
               if str(row[clmn]) == "nan" or str(row[clmn].strip()) == " ":
                  subCount=subCount+1
                  continue
               else:
                  #print("ColCount : ", str(colCount-1))
                  print("{}{}{}".format(dqt,str(row[clmn].strip()),dqt), end="" )
                  subCount=subCount+1
            else:
               if str(row[clmn]) == "nan" or str(row[clmn]) == " ":
                  subCount=subCount+1
                  continue
               else:
                  #print("\"{}\",".format(str(row[clmn].strip())), end="" )
                  print("{}{}{}{}".format(dqt,str(row[clmn].strip()),dqt,cma), end="" )
                  subCount=subCount+1
       if k == len(df3) - 1: 							 #If last row, will not add comma, else will add a comma along with the close square bracket
          print(" {}\n".format(csq))
       else:
          print(" {}{}\n".format(csq,cma))

print(" {}".format(ccrl))

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#subnet vnet mapping
#
#below to getrid of NaN
#df4['vnet name'].replace(' ', pd.np.nan, inplace=True)
df4['vnet name'].replace(' ', nump.nan, inplace=True)
df4.dropna(subset=['vnet name'], inplace=True)
df4=df4.fillna('')
print("\n\n#Input to create the subnets" )
print("\n\n\"subnet_vnet_mapping\" = \n[" )
for l, (index, row) in enumerate(df4.iterrows()):
    if l == len(df4) - 1:
       print("\t{ subnet_name=\"" + row['subnet name'].strip() + "\" rg_name=\"" + row['resource group'].strip() + "\" vnet_name=\"" + row['vnet name'].strip() + "\" subnet_address_prefix=\"" + row['address space'].strip() + "\" }" )
    else:
       print("\t{ subnet_name=\"" + row['subnet name'].strip() + "\" rg_name=\"" + row['resource group'].strip() + "\" vnet_name=\"" + row['vnet name'].strip() + "\" subnet_address_prefix=\"" + row['address space'].strip() + "\" }," )
print("]" )


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#Public Ip addresses
#
#below to getrid of NaN
#df5['public ip name'].replace(' ', pd.np.nan, inplace=True)
df5['public ip name'].replace(' ', nump.nan, inplace=True)
df5.dropna(subset=['public ip name'], inplace=True)
df5=df5.fillna('')
print("\n\n#Input to create the public ip address" )
print("\n\n\"public_ip_address\" = \n[" )
for m, (index, row) in enumerate(df5.iterrows()):
    if m == len(df5) - 1:
       print("\t{ public_ip_name=\"" + row['public ip name'].strip() + "\" rg_name=\"" + row['public resource group'].strip() + "\" location=\"" + row['location'].strip() + "\" vnet_name=\"" + row['vnet'].strip() + "\" allocation_method=\"" + row['static'].strip() + "\" sku=\"" + row['sku'].strip() + "\" }" )
    else:
       print("\t{ public_ip_name=\"" + row['public ip name'].strip() + "\" rg_name=\"" + row['public resource group'].strip() + "\" location=\"" + row['location'].strip() + "\" vnet_name=\"" + row['vnet'].strip() + "\" allocation_method=\"" + row['static'].strip() + "\" sku=\"" + row['sku'].strip() + "\" }," )
print("]" )

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#Network Security groups (NSG)
#
#below to getrid of NaN
#df6['nsg-name'].replace(' ', pd.np.nan, inplace=True)
df6['nsg-name'].replace(' ', nump.nan, inplace=True)
df6.dropna(subset=['nsg-name'], inplace=True)
df6=df6.fillna('')
print("\n\n#Input to create the NSG" )
print("\n\n\"network_security_group\" = \n[", )
for n, (index, row) in enumerate(df6.iterrows()):
    if n == len(df6) - 1:
       print("\t{ nsg_name=\"" + row['nsg-name'].strip() + "\" location=\"" + row['location'].strip() + "\" rg_name=\"" + row['resource-group'].strip() + "\" }" )
    else:
       print("\t{ nsg_name=\"" + row['nsg-name'].strip() + "\" location=\"" + row['location'].strip() + "\" rg_name=\"" + row['resource-group'].strip() + "\" }," )
print("]" )

#azurerm_network_security_rule" = (E through N on NSG tab)
#Added March-12-2020
#Rule inbound/outbound,priority,name,port,protocol,source,destination,action,nic,subnet
#nsg_name="" rg_name="" rule_name="" inoutbound="" key="" prio="" sr_port="" dest="" proto="tcp" src_add="*" dst_add="*"

print("\n\n\"azurerm_network_security_rule\" = \n[", )
for n, (index, row) in enumerate(df6.iterrows()):
    if n == len(df6) - 1:
       print("\t{ nsg_name=\"" + row['nsg-name'].strip() + "\" rg_name=\"" + row['resource-group'].strip() 
       + "\" rule_name=\"" + row['Rule inbound/outbound'].strip() + "\" key=\"" + row['action'].strip() 
       + "\" prio=\"" + str(row['priority']) + "\" sr_port=\"" + row['port'] + "\" dest=\"" + row['destination'].strip() 
       + "\" proto=\"" + row['protocol'] + "\" src_addr=\"" + row['source'].strip() + "\" dst_addr=\"" + row['destination'].strip()  + "\" }" )
    else:
       print("\t{ nsg_name=\"" + row['nsg-name'].strip() + "\" rg_name=\"" + row['resource-group'].strip() 
       + "\" rule_name=\"" + row['Rule inbound/outbound'].strip() + "\" key=\"" + row['action'].strip() 
       + "\" prio=\"" + str(row['priority']) + "\" sr_port=\"" + row['port'] + "\" dest=\"" + row['destination'].strip() 
       + "\" proto=\"" + row['protocol'] + "\" src_addr=\"" + row['source'].strip() + "\" dst_addr=\"" + row['destination'].strip()  + "\" }," )
print("]" )


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#UDR Route tables
#
#below to getrid of NaN
#df7['rg_name'].replace(' ', pd.np.nan, inplace=True)
df7['rg_name'].replace(' ', nump.nan, inplace=True)
df7.dropna(subset=['rg_name'], inplace=True)
df7=df7.fillna('')
print("\n\n#Input to create UDR Route tables" )
print("\n\n\"route_table\" = \n[", )
for m, (index, row) in enumerate(df7.iterrows()):
    if m == len(df7) - 1:
       print("\t{ rg_name=\"" + row['rg_name'].strip() + "\" udr_name=\"" + row['udr_name'].strip() + "\" location=\"" + row['location'].strip() + "\" }" )
    else:
       #print >> fp, "\tnsg_name=\"" + row['nsg-name'].strip() + "\" location=\"" + row['location'].strip() + "\" rg_name=\"" + row['resource-group'].strip() + "\" },"
       print("\t{ rg_name=\"" + row['rg_name'].strip() + "\" udr_name=\"" + row['udr_name'].strip() + "\" location=\"" + row['location'].strip() + "\" }," )
print("]" )

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#Vnet peering
#
#vnetpeer1 = {
#    peer_name = "dc46-app-peer1"
#    rg_name = "dc46-network-rg"
#    vnet_name = "dc46-app-vnet"
#    remote_virtual_network_id = "test2.id" }
#below to getrid of NaN
df8=df8.fillna('')
print("\n\n#Input to create Vnet peering" )
for o, (index, row) in enumerate(df8.iterrows()):
       #Below pre network-tffile_v1.1.xlsx 
       #print("\n" + row['Peernum'].strip() + " = {" + "\n\tpeer_name=\"" + row['peer_name local'].strip() + "\" \n\trg_name=\"" + row['rg_name'].strip() + "\" \n\tvnet_name=\"" + row['vnet_name'].strip() + "\" \n\tremote_virtual_network_id=\"" + row['remote_virtual_network_id'] + "\" \n }" )
       if row['peer_name local'].strip() == "":
          continue
       else :
       	  print("\n" + row['peer_name local'].strip() + " = {" + "\n\tpeer_name=\"" + row['peer_name local'].strip() + "\" \n\trg_name=\"" + row['rg_name'].strip() + "\" \n\tvnet_name=\"" + row['vnet_name'].strip() + "\" \n }" )

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
#Close the open file, reset stdout to screen
#
fp.close()
#below to reset standard out to screen
sys.stdout=sys.__stdout__
print("Done writing to terraform.tfvars")

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#
# New requirement : check the difference of previous terraform.tfvars.timestamp with new terraform.tfvars
#
#if bkp:
#   print("\n\nPrevious terraform.tfvars exists below is the difference between previous and the newely Generated terraform.tfvars\n\n")
#   with open("terraform.tfvars") as Curr:
#        CurrentFile=Curr.readlines()
#   with open(bkp) as Previous:
#        PreviousFile=Previous.readlines()
#   #Find and print the diff, n=0 else by default will display 3 context lines
#   for line in difflib.unified_diff(CurrentFile, PreviousFile, fromfile=Curr.name, tofile=Previous.name, n=0, lineterm=""):
#       print(line)
#   #for line in difflib.unified_diff(CurrentFile, PreviousFile):
#   #    print(line)
#   #diff=difflib.ndiff(CurrentFile, PreviousFile)
#   #diff=difflib.SequenceMatcher(None,CurrentFile,PreviousFile)
#   #print(diff),
#   #print(''.join(diff)),

#   Curr.close()
#   Previous.close()
print('This run of ' + scriptName + ' script took {0} second'.format(time.time() - startTime))
#End of Script
