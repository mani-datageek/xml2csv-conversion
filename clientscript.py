import xml.etree.ElementTree as ET
import pandas as pd
import datetime
import pytz
import os

# Define the date conversion function
def conv_datetime_to_Pacific(iso_date_string):
    date_object = datetime.datetime.strptime(iso_date_string, "%Y-%m-%dT%H:%M:%S%z")
    pacific_timezone = pytz.timezone("US/Pacific")
    pacific_date = date_object.astimezone(pacific_timezone)
    pacific_date_string = pacific_date.strftime("%Y-%m-%d %H:%M:%S")
    return pacific_date_string

# Folder path containing XML files
folder_path = '/Users/manisharma/Documents/'

# Lists to store results
Result = []
Associations = []

# Loop through each XML file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".xml"):
        file_path = os.path.join(folder_path, filename)
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Iterate through each helpdesk-ticket element
        for ticket in root.findall('helpdesk-ticket'):
            ticket_number = ticket.findtext(".//display-id")
            description = ticket.findtext(".//description")
            ticket_id = ticket.findtext(".//id")
            requestor = ticket.findtext(".//requester-name")
            requestor_id = ticket.findtext(".//requester-id")
            ticket_created = ticket.findtext(".//created-at")
            ticket_created = conv_datetime_to_Pacific(ticket_created)
            
            # Save the description as the first communication of the ticket
            note_data = {
                "Ticket_Number": ticket_number,
                "Response_Date": ticket_created,
                "Responder": requestor_id,
                "Note_ID": ticket_id,
                "Description": description,
            }
            Result.append(note_data)
            
            #---Section related to ticket association fields---
            association_type_string = ticket.findtext(".//association-type")
            if not association_type_string:
                association_type = 0
            else:
                association_type = int(association_type_string)
            association_rdb = ticket.findtext(".//associates-rdb")
            if association_type == 1:
                association_name = "Parent"
            elif association_type == 2:
                association_name = "Child"
            else:
                association_name = "None"
            association_data = {
                "Ticket_Number": ticket_number,
                "Association_Type": association_name,
                "Associates_RDB": association_rdb,
            }
            Associations.append(association_data)
            
            # Iterate through each note item
            for note in ticket.findall(".//helpdesk-note"):
                response_date = note.findtext("created-at")
                response_date = conv_datetime_to_Pacific(response_date)
                note_id = note.findtext("id")
                note_description = note.findtext("body")
                user_id = note.findtext(".//user-id")
                note_data = {
                    "Ticket_Number": ticket_number,
                    "Response_Date": response_date,
                    "Responder": user_id,
                    "Note_ID": note_id,
                    "Description": note_description,
                }
                Result.append(note_data)

# Create DataFrames
df_result = pd.DataFrame(Result)
df_associations = pd.DataFrame(Associations)

# Export to Excel
df_result.to_excel("output_tickets.xlsx", index=False)
df_associations.to_excel("output_associations.xlsx", index=False)
