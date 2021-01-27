import json
import requests
import csv
import io
import time

print("starting...")
###########################################################################
# Name: Hubspot Spreadsheet Export                                        #
# Description: This script let's you connect to Hubspot CRM and retrieve  #
#              its data to populate a csv file.                           #
# Author: Ryan Wong, Luca Perrone                                         #
###########################################################################

###########################################################################
# ----------------------------------------------------------------------- #
# ------------------------------- CONFIG -------------------------------- #
# ----------------------------------------------------------------------- #
###########################################################################

# import json file containing the api key
with open("C:/Sales data/secret_keys.json") as f:
  data = json.load(f)

#set base variables
start_time = time.time()
SCOPE = 'contacts'
API_KEY = data["hubspot_api_key"]
API_URL_BASE = 'https://api.hubapi.com'


###########################################################################
# ----------------------------------------------------------------------- #
# ------------------------------- GET DATA ------------------------------ #
# ----------------------------------------------------------------------- #
###########################################################################

# GET COMPANY
print ("Retrieving Company Data")
# set the base url and properties to include in company table
url = "https://api.hubapi.com/crm/v3/objects/companies"

properties = "name,industry,market_segment_company,city,state,country,relationship_manager,hs_object_id"

# set list that will contain all company in Flinks account and last page status for loop
companies = []
last_page = False

# make the first call to GET company endpoint and store the paging cursor token for loop
querystring = {"limit":"100","properties":properties,
"paginateAssociations":"false","archived":"false","hapikey":API_KEY}

headers = {'accept': 'application/json'}

response = requests.request("GET", url, headers=headers, params=querystring)

data = json.loads(response.text)

company_data = data['results']

companies.extend(company_data)

after = data['paging']['next']['after']

# loop through pages of GET company calls of 100 company 
while not last_page:
    querystring = {"limit":"100","after":after,"properties":properties,
    "paginateAssociations":"false","archived":"false","hapikey":API_KEY}

    headers = {'accept': 'application/json'}

    response = requests.request("GET", url, headers=headers, params=querystring)

    data = json.loads(response.text)

    company_data = data['results']

    if len(company_data) != 100:
        last_page = True

    companies.extend(company_data)

    if not last_page:
        after = data['paging']['next']['after']

print("Companies (All companies): " + str(len(companies)))

###########################################################################
# ----------------------------------------------------------------------- #
# -------------------------- WRITE TO SPREADSHEET ----------------------- #
# ----------------------------------------------------------------------- #
###########################################################################

# WRITE COMPANY TO SPREADSHEET/CSV

# now we will open a file for writing 
company_file = io.open('C:/Sales data/CSVs/company_file.csv', 'w', encoding="utf-8")

# create the csv writer object 
csv_writer = csv.writer(company_file)
  
# Boolean variable used for writing  
# headers to the CSV file 
titles = True

for company in companies: 
    
    company_dict = {
        "Name": company['properties']['name'],
        "Industry": company['properties']['industry'],
        "Market Segment": company['properties']['market_segment_company'],
        "City": company['properties']['city'],
        "State/region": company['properties']['state'],
        "Country/Region": company['properties']['country'],
        "Relationship Manager": company['properties']['relationship_manager'],
        "Company ID":company['properties']['hs_object_id']
    }
    
    if titles: 

        # Writing headers of CSV file 
        header = company_dict.keys()
        csv_writer.writerow(header)
        titles = False
    
    # Writing data of CSV file 
    csv_writer.writerow(company_dict.values())
  
company_file.close()
print("The Spreadsheet Took --- %s seconds --- To be Generated" % (time.time() - start_time))