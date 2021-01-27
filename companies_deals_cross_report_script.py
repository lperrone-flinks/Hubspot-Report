import json
import requests
import csv
import io
import time

print("starting...")
print("Retrieving Cross Report of Companies and Deals")
###########################################################################
# Name: Hubspot Spreadsheet Export                                        #
# Description: This script let's you connect to Hubspot CRM and retrieve  #
#              its data to populate a csv file.                           #
# Author: Ryan Wong, Luca Perrone                                         #
###########################################################################
# import json file containing the api key
with open("C:/Sales data/secret_keys.json") as f:
  data = json.load(f)


start_time = time.time()
API_KEY = data["hubspot_api_key"]

###########################################################################
# ----------------------------------------------------------------------- #
# ------------------------------- GET DATA ------------------------------ #
# ----------------------------------------------------------------------- #
###########################################################################

# set the base url and properties to include in deals table
deal_url = "https://api.hubapi.com/crm/v3/objects/deals"

deal_properties = "amount_in_home_currency,weighted_company_amount,dealstage,dealname,client_size,hubspot_owner_id,closedate,became_closed_onboarding_date,live_date,became_live_date,billed,account_manager,pipeline"

# set list that will contain all deals in Flinks account and last page status for loop
deals = []
last_page = False

# make the first call to GET deals endpoint and store the paging cursor token for loop
deal_querystring = {"limit":"100","properties":deal_properties,"associations":"Company",
"paginateAssociations":"false","archived":"false","hapikey":API_KEY}

headers = {'accept': 'application/json'}

deal_response = requests.request("GET", deal_url, headers=headers, params=deal_querystring)

d_data = json.loads(deal_response.text)

deal_data = d_data['results']

deals.extend(deal_data)

after = d_data['paging']['next']['after']

# loop through pages of GET deals calls of 100 deals (ie 51 loops to retrieve 5080 deals)
while not last_page:
    deal_querystring = {"limit":"100","after":after,"properties":deal_properties,"associations":"Company",
    "paginateAssociations":"false","archived":"false","hapikey":API_KEY}

    headers = {'accept': 'application/json'}

    deal_response = requests.request("GET", deal_url, headers=headers, params=deal_querystring)

    d_data = json.loads(deal_response.text)

    deal_data = d_data['results']

    if len(deal_data) != 100:
        last_page = True

    deals.extend(deal_data)

    if not last_page:
        after = d_data['paging']['next']['after']

# loop through and reformat deals with correct header names into deals_list
deals_list = []

for deal in deals: 
    
    associated_company_ids = []
    
    if 'associations' in deal:
        for result in deal['associations']['companies']['results']:
            associated_company_ids.append(result['id'])

    deal_dict = {
        "Amount in company currency": deal['properties']['amount_in_home_currency'],
        "weighted Company Amount": deal['properties']['weighted_company_amount'],
        "Deal Stage": deal['properties']['dealstage'],
        "Deal Name": deal['properties']['dealname'],
        "Company Persona": deal['properties']['client_size'],
        "Deal owner": deal['properties']['hubspot_owner_id'],
        "Close date": deal['properties']['closedate'],
        "Became Closed - Onboarding Date": deal['properties']['became_closed_onboarding_date'],
        "Live Date": deal['properties']['live_date'],
        "Became live date": deal['properties']['became_live_date'],
        "Billing Date": deal['properties']['billed'],
        "Relationship Manager": deal['properties']['account_manager'],
        "Associated Company ID": associated_company_ids,
        "Pipeline": deal['properties']['pipeline']
    }

    deals_list.append(deal_dict)

print("Deals (All pipelines): " + str(len(deals_list)))

# create a deals dictionary with key = company IDs, and values = list of deal objects
deals_dict = {}

for deal in deals_list:
    companyIds = deal["Associated Company ID"] 
    for companyId in companyIds:
        if deals_dict.get(companyId) is None:
            deals_dict.update({companyId: [deal]})
        else:
            deals = deals_dict.get(companyId)
            deals.append(deal)
            deals_dict.update({companyId: deals})

print("Dictionary of Deal IDs: " + str(len(deals_dict)))

companies_list = []

# opening the companies file into a list of JSON objects 
companies_data = io.open('C:/Sales data/CSVs/company_file.csv', 'r', encoding="utf-8")
      
for line in csv.DictReader(companies_data): 
    companies_list.append(line)

print("Companies (All companies): " + str(len(companies_list)))

###########################################################################
# ----------------------------------------------------------------------- #
# -------------------------- WRITE TO SPREADSHEET ----------------------- #
# ----------------------------------------------------------------------- #
###########################################################################

# WRITE CROSS REPORT TO SPREADSHEET/CSV

# now we will open a file for writing 
cross_report_file = io.open('C:/Sales data/CSVs/company_deal_cross_report.csv', 'w', encoding="utf-8")

# create the csv writer object 
csv_writer = csv.writer(cross_report_file)

# Boolean variable used for writing  
# headers to the CSV file 
titles = True

# Counter variable used to check if
# there is a first valid entry into file
firstValidEntry = 0

# Counter variable for entries in cross report
count = 0

for company in companies_list:
    reports = []
    if company["Company ID"] in deals_dict:
        for deal in deals_dict[company["Company ID"]]:
            if deal["Deal Name"].lower().startswith('id') or deal["Deal Name"].lower().startswith('doordash') or deal["Deal Name"].lower().startswith('goeasy') or deal["Deal Name"].lower().startswith('qid'):
                firstValidEntry += 1
                report = {}
                report.update(company)
                report.update(deal)
                reports.append(report)
    
    if titles and firstValidEntry == 1: 

        # Writing headers of CSV file 
        header = reports[0].keys()
        csv_writer.writerow(header)
        titles = False
    
    # Writing data of CSV file 
    for report in reports:
        csv_writer.writerow(report.values())
        count += 1

cross_report_file.close()

print("Cross report records: " + str(count))

print("The Spreadsheet Took --- %s seconds --- To be Generated" % (time.time() - start_time))