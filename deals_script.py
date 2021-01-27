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

# GET DEALS
print ("Retrieving Deals Data")
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

print("Deals (All pipelines): " + str(len(deals)))

# make a list that will store only deals from Flinks Sales Pipeline (ie 1837)
sales_pipeline_deals = []

for deal in deals:
    if (deal['properties']['pipeline'] == "1208785"):
        sales_pipeline_deals.append(deal)

print("Deals (Flinks Sales Pipeline): " + str(len(sales_pipeline_deals)))

###########################################################################
# ----------------------------------------------------------------------- #
# -------------------------- WRITE TO SPREADSHEET ----------------------- #
# ----------------------------------------------------------------------- #
###########################################################################

# WRITE DEALS TO SPREADSHEET/CSV

# now we will open a file for writing 
deal_file = io.open('C:/Sales data/CSVs/deal_file.csv', 'w', encoding="utf-8")

# create the csv writer object 
csv_writer = csv.writer(deal_file)
  
# Boolean variable used for writing  
# headers to the CSV file 
titles = True

for deal in sales_pipeline_deals: 
    
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
    
    if titles: 

        # Writing headers of CSV file 
        header = deal_dict.keys()
        csv_writer.writerow(header)
        titles = False
    
    # Writing data of CSV file 
    csv_writer.writerow(deal_dict.values())
  
deal_file.close()
print("The Spreadsheet Took --- %s seconds --- To be Generated" % (time.time() - start_time))