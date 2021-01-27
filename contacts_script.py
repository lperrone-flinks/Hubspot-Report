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

# GET CONTACT
print ("Retrieving Contact Data")
# set the base url and properties to include in contacts table
url = "https://api.hubapi.com/crm/v3/objects/contacts"


#add in properties 

properties = "company,firstname,lastname,email,phone,hubspot_owner_id,notes_last_updated,hs_lead_status,createdate,hs_analytics_source,hs_email_domain,hs_lifecyclestage_salesqualifiedlead_date,hs_analytics_num_visits,hs_analytics_last_touch_converting_campaign,hs_analytics_first_touch_converting_campaign,num_conversion_events,lifecyclestage,recent_deal_amount,recent_conversion_event_name,associatedcompanyid"

# set list that will contain all contact in Flinks account and last page status for loop
contacts = []
last_page = False

# make the first call to GET contact endpoint and store the paging cursor token for loop
querystring = {"limit":"100","properties":properties,
"paginateAssociations":"false","archived":"false","hapikey":API_KEY}

headers = {'accept': 'application/json'}

response = requests.request("GET", url, headers=headers, params=querystring)

data = json.loads(response.text)

contact_data = data['results']

contacts.extend(contact_data)

after = data['paging']['next']['after']

# loop through pages of GET contact calls of 100 contact 
while not last_page:
    querystring = {"limit":"100","after":after,"properties":properties,
    "paginateAssociations":"false","archived":"false","hapikey":API_KEY}

    headers = {'accept': 'application/json'}

    response = requests.request("GET", url, headers=headers, params=querystring)

    data = json.loads(response.text)

    contact_data = data['results']

    if len(contact_data) != 100:
        last_page = True

    contacts.extend(contact_data)

    if not last_page:
        after = data['paging']['next']['after']

print("Contacts (All contacts): " + str(len(contacts)))

###########################################################################
# ----------------------------------------------------------------------- #
# -------------------------- WRITE TO SPREADSHEET ----------------------- #
# ----------------------------------------------------------------------- #
###########################################################################

# WRITE contact TO SPREADSHEET/CSV

# now we will open a file for writing 
contact_file = io.open('C:/Sales data/CSVs/contact_file.csv', 'w', encoding="utf-8")

# create the csv writer object 
csv_writer = csv.writer(contact_file)
  
# Boolean variable used for writing  
# headers to the CSV file 
titles = True

for contact in contacts: 
    
    contact_dict = {
        "Contact ID": contact['properties']['hs_object_id'],
        "First Name": contact['properties']['firstname'],
        "Last Name": contact['properties']['lastname'],
        "Email": contact['properties']['email'],
        "Phone Number": contact['properties']['phone'],
        "Contact Owner": contact['properties']['hubspot_owner_id'],
        "Last Activity Date": contact['properties']['notes_last_updated'],
        "Lead Status": contact['properties']['hs_lead_status'],
        "Create Date": contact['properties']['createdate'],
        "Origional Source": contact['properties']['hs_analytics_source'],
        "Email Domain": contact['properties']['hs_email_domain'],
        "Became a Sales Qualified Lead Date": contact['properties']['hs_lifecyclestage_salesqualifiedlead_date'],
        "Number of Sessions": contact['properties']['hs_analytics_num_visits'],
        "Last Touch Converting Campaign": contact['properties']['hs_analytics_last_touch_converting_campaign'],
        "First Touch Converting Campaign": contact['properties']['hs_analytics_first_touch_converting_campaign'],
        "Number of Form Submissions": contact['properties']['num_conversion_events'],
        "Lifecycle Stage": contact['properties']['lifecyclestage'],
        "Recent Deal Amount": contact['properties']['recent_deal_amount'],
        "Recent Conversion": contact['properties']['recent_conversion_event_name'],
        "Associated Company ID": contact['properties']['associatedcompanyid'],
    }
    
    if titles: 

        # Writing headers of CSV file 
        header = contact_dict.keys()
        csv_writer.writerow(header)
        titles = False
    
    # Writing data of CSV file 
    csv_writer.writerow(contact_dict.values())
  
contact_file.close()
print("The Spreadsheat Took --- %s seconds --- To be Generated" % (time.time() - start_time))