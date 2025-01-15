import requests
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

# Calculate yesterday's date
yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

# Construct the RSS feed URL
rss_url = f"https://clinicaltrials.gov/api/rss?firstPost={yesterday}_{yesterday}&dateField=StudyFirstPostDate"

# Fetch the RSS feed
response = requests.get(rss_url)

# Check if the request was successful
if response.status_code == 200:
    # Parse the XML response
    root = ET.fromstring(response.content)

    # Extract study IDs
    study_ids = []
    for item in root.findall('.//item'):
        link = item.find('link').text
        if "study/" in link:
            # Extract the study ID from the link
            study_id = link.split("study/")[1].split("?")[0]
            study_ids.append(study_id)

    print(f"\nTotal count of study IDs: {len(study_ids)}")
    
    # Define the API endpoint for fetching study information
    base_url = "https://clinicaltrials.gov/api/v2/studies"
    results = []

    # Function to fetch study information for a chunk of IDs
    def fetch_study_info(base_url, ids):
        url = f"{base_url}?filter.ids={ids}"
        try:
            response = requests.get(url)
            response.raise_for_status()  # Raise an error for bad responses
            data = response.json()
            return data['studies']  # Return the list of studies
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data: {e}")
            return []

    # Chunk size
    chunk_size = 100
    for i in range(0, len(study_ids), chunk_size):
        chunk = study_ids[i:i + chunk_size]  # Get a chunk of IDs
        ids = '|'.join(chunk)  # Create a pipe-separated string
        studies = fetch_study_info(base_url, ids)
        results.extend(studies)  # Add the fetched studies to the results

    # Check if we need to handle pagination
    total_results = len(results)
    while total_results < len(study_ids):
        # Fetch more studies if available
        next_chunk = study_ids[total_results:total_results + chunk_size]
        if not next_chunk:
            break
        ids = '|'.join(next_chunk)
        studies = fetch_study_info(base_url, ids)
        results.extend(studies)
        total_results = len(results)

    # Extract specific fields
    extracted_data = []
    for study in results:
        if 'protocolSection' in study:
            protocol = study['protocolSection']
            sponsor_info = protocol.get('sponsorCollaboratorsModule', {}).get('leadSponsor', {})
            contacts_info = protocol.get('contactsLocationsModule', {}).get('centralContacts', [])
            brief_title = protocol.get('identificationModule', {}).get('briefTitle', "N/A")  # Get the Brief Title
            
            # Prepare the extracted information
            if contacts_info:
                for contact in contacts_info:
                    extracted_info = {
                        "brief_title": brief_title,
                        "lead_sponsor": sponsor_info.get("name", "N/A"),
                        "contact_name": contact.get("name", "N/A"),
                        "contact_phone": contact.get("phone", "N/A"),
                        "contact_email": contact.get("email", "N/A")
                    }
                    extracted_data.append(extracted_info)
            else:
                # If there are no contacts, still add the study with empty contact info
                extracted_info = {
                    "brief_title": brief_title,
                    "lead_sponsor": sponsor_info.get("name", "N/A"),
                    "contact_name": "N/A",
                    "contact_phone": "N/A",
                    "contact_email": "N/A"
                }
                extracted_data.append(extracted_info)

    # Save extracted data to XLSX file
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Clinical Trials Data"

    # Add headers
    headers = ["Brief Title", "Lead Sponsor", "Contact Name", "Contact Phone", "Contact Email"]
    sheet.append(headers)

    # Add data rows
    for data in extracted_data:
        row = [
            data["brief_title"],
            data["lead_sponsor"],
            data["contact_name"],
            data["contact_phone"],
            data["contact_email"]
        ]
        sheet.append(row)

    # Format the header row
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Create a table
    table = Table(displayName="ClinicalTrialsTable", ref=f"A1:E{len(extracted_data) + 1}")

    # Add a table style without highlighting the first column
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style

    # Add the table to the sheet
    sheet.add_table(table)

    # Adjust column widths
    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]  # Convert to list to access headers and cells
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)  # Add a little extra space
        sheet.column_dimensions[column[0].column_letter].width = adjusted_width

    # Save the workbook
    output_filename = f"clinical_trials_data_{yesterday}.xlsx"
    workbook.save(output_filename)
    print(f"Data saved to {output_filename}")

else:
    print("Failed to fetch data. Status code:", response.status_code)