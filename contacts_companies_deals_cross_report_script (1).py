import json
import requests
import csv
import io
import time

print("starting...")
print("Retrieving Cross Report of Contacts, Companies and Deals")
###########################################################################
# Name: Hubspot Spreadsheet Export                                        #
# Description: This script let's you connect to Hubspot CRM and retrieve  #
#              its data to populate a csv file.                           #
# Author: Ryan Wong, Luca Perrone                                         #
###########################################################################

start_time = time.time()

###########################################################################
# ----------------------------------------------------------------------- #
# ------------------------------- GET DATA ------------------------------ #
# ----------------------------------------------------------------------- #
###########################################################################

company_deal_cross_report_list = []

# opening the companies and deals cross report file into a list of JSON objects 
cross_report_data = io.open('C:/Sales data/CSVs/company_deal_cross_report.csv', 'r', encoding="utf-8")
      
for line in csv.DictReader(cross_report_data): 
    company_deal_cross_report_list.append(line)

print("Companies/Deals cross report records: " + str(len(company_deal_cross_report_list)))

contacts_list = []

# opening the contacts file into a list of JSON objects 
contacts_data = io.open('C:/Sales data/CSVs/contact_file.csv', 'r', encoding="utf-8")
      
for line in csv.DictReader(contacts_data): 
    contacts_list.append(line)

print("Contacts (All contacts): " + str(len(contacts_list)))

###########################################################################
# ----------------------------------------------------------------------- #
# -------------------------- WRITE TO SPREADSHEET ----------------------- #
# ----------------------------------------------------------------------- #
###########################################################################

# WRITE CROSS REPORT TO SPREADSHEET/CSV

# now we will open a file for writing 
cross_report_file = io.open('C:/Sales data/CSVs/contacts_cross_report.csv', 'w', encoding="utf-8")

# create the csv writer object 
csv_writer = csv.writer(cross_report_file)

# Boolean variable used for writing  
# headers to the CSV file 
titles = True

# Counter variable for entries in cross report
count = 0

for contact in contacts_list:
    report = {}
    for row in company_deal_cross_report_list:
        if contact["Associated Company ID"] == row["Company ID"]:
            report.update(contact)
            report.update(row)
        else:
            report.update(contact)
    
    if titles: 

        # Writing headers of CSV file 
        #header = report.keys()
        csv_writer.writerow(['Contact ID', 'First Name', 'Last Name', 'Email', 'Phone Number', 'Contact Owner', 'Last Activity Date', 'Lead Status', 'Create Date', 'Origional Source', 'Email Domain', 'Became a Sales Qualified Lead Date', 'Number of Sessions', 'Last Touch Converting Campaign', 'First Touch Converting Campaign', 'Number of Form Submissions', 'Lifecycle Stage', 'Recent Deal Amount', 'Recent Conversion', 'Associated Company ID', 'Name', 'Industry', 'Market Segment', 'City', 'State/region', 'Country/Region', 'Relationship Manager', 'Company ID', 'Amount in company currency', 'weighted Company Amount', 'Deal Stage', 'Deal Name', 'Company Persona', 'Deal owner', 'Close date', 'Became Closed - Onboarding Date', 'Live Date', 'Became live date', 'Billing Date', 'Pipeline'])
        titles = False
    
    # Writing data of CSV file 
    csv_writer.writerow(report.values())
    count += 1

cross_report_file.close()

print("Contacts cross report records: " + str(count))

print("The Spreadsheet Took --- %s seconds --- To be Generated" % (time.time() - start_time))