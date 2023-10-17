import requests
import urllib.parse
import openpyxl
import csv
import json
from datetime import datetime, timedelta

# Replace 'YOUR_ACCESS_TOKEN' with your authentication token
ACCESS_TOKEN = 'YOUR_ACCESS_TOKEN'

# API URL
API_BASE_URL = 'https://webexapis.com/v1'

# Common headers for API requests
HEADERS = {
    'Authorization': f'Bearer {ACCESS_TOKEN}',
    'Content-Type': 'application/json',
}

# Default aggregation and maximum time spans
DEFAULT_AGGREGATION = 'hourly'
HOURLY_MAX_TIME_SPAN = 47
DAILY_MAX_TIME_SPAN = 29

# List of possible metric names
METRIC_NAMES = ['duration', 'soundLevel', 'ambientNoise', 'temperature', 'humidity', 'tvoc', 'peopleCount']

# Function to make a GET request to the Webex API
def api_get_request(endpoint, params=None):
    response = requests.get(endpoint, headers=HEADERS, params=params)
    return response

# Function to check the access token
def check_access_token(access_token, headers):
    if access_token == 'YOUR_ACCESS_TOKEN':
        access_token = input("\n\033[0;38mEnter your Webex API access token: ")

    headers['Authorization'] = f'Bearer {access_token}'  # Update headers
    api_check_url = f'{API_BASE_URL}/people/me'
    response = api_get_request(api_check_url, headers)

    if response.status_code == 200:
        return f"\n\033[0;32mToken access correct\n"
    else:
        print("\n\033[0;31mToken access failed. Please check your access token.")
        access_token = input("\n\033[0;38mEnter your Webex API access token: ")  # Ask the user to enter the correct access token
        return check_access_token(access_token, headers)

# Function to get locationId using the location name
def get_location_id(location_name):
    encoded_location_name = urllib.parse.quote(location_name)
    request_url = f'{API_BASE_URL}/workspaceLocations?displayName={encoded_location_name}'
    response = api_get_request(request_url)

    if response.status_code == 200:
        data = response.json()
        if data.get('items'):
            return data['items'][0]['id']
    return None

# Function to get floorId using locationId
def get_floor_id(location_id):
    request_url = f'{API_BASE_URL}/workspaceLocations/{location_id}/floors'
    response = api_get_request(request_url)

    if response.status_code == 200:
        data = response.json()
        if data.get('items'):
            return data['items']  # Return a list of floor objects
    return []

# Function to get workspaceId using locationId and floorId
def get_workspace_id(location_id, floor_id):
    request_url = f'{API_BASE_URL}/workspaces'
    params = {
        'locationId': location_id,
        'floorId': floor_id
    }
    response = api_get_request(request_url, params)

    if response.status_code == 200:
        data = response.json()
        if data.get('items'):
            return data['items']  # Return a list of workspace objects
    return []

# Function to get workspace name using workspaceId
def get_workspace_name(workspace_id):
    request_url = f'{API_BASE_URL}/workspaces/{workspace_id}'
    response = api_get_request(request_url)

    if response.status_code == 200:
        data = response.json()
        return data.get('displayName')  # Get the workspace name from the API response
    return None

# Function to get workspace metrics using workspaceId
def get_workspace_metrics(workspace_id, metric_name, aggregation, from_date_time_iso, to_date_time_iso):
    encoded_workspace_id = urllib.parse.quote(workspace_id)
    encoded_metric_name = urllib.parse.quote(metric_name)
    encoded_aggregation = urllib.parse.quote(aggregation)
    encoded_from_date_time_iso = urllib.parse.quote(from_date_time_iso)
    encoded_to_date_time_iso = urllib.parse.quote(to_date_time_iso)

    request_url = f'{API_BASE_URL}/workspaceMetrics?workspaceId={encoded_workspace_id}&metricName={encoded_metric_name}&aggregation={encoded_aggregation}&from={encoded_from_date_time_iso}Z&to={encoded_to_date_time_iso}Z'
    response = api_get_request(request_url)

    if response.status_code == 200:
        data = response.json()
        return data
    return None

# Function to get workspace duration metrics using workspaceId
def get_workspace_duration_metrics(workspace_id, aggregation, from_date_time_iso, to_date_time_iso):
    encoded_workspace_id = urllib.parse.quote(workspace_id)
    encoded_aggregation = urllib.parse.quote(aggregation)
    encoded_from_date_time_iso = urllib.parse.quote(from_date_time_iso)
    encoded_to_date_time_iso = urllib.parse.quote(to_date_time_iso)

    request_url = f'{API_BASE_URL}/workspaceDurationMetrics?workspaceId={encoded_workspace_id}&aggregation={encoded_aggregation}&from={encoded_from_date_time_iso}Z&to={encoded_to_date_time_iso}Z'
    response = api_get_request(request_url)

    if response.status_code == 200:
        data = response.json()
        return data
    return None

# Function to get workspace capacity using workspaceId
def get_workspace_capacity(workspace_id):
    request_url = f'{API_BASE_URL}/workspaces/{workspace_id}'
    response = api_get_request(request_url)

    if response.status_code == 200:
        data = response.json()
        return data.get('capacity', 'N/A')
    return 'N/A'

# Function to export data in various formats
def export_data(data, headers, filename, export_format):
    if export_format == 'xlsx':
        # Create an XLSX file
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = 'Workspace Metrics'

        # Add headers
        worksheet.append(headers)

        # Add data to the worksheet
        for data_row in data:
            worksheet.append([data_row[key] for key in headers])

        # Save the XLSX file
        workbook.save(filename)
        print(f"\033[0;32mData saved to {filename}")

    elif export_format == 'csv':
        # Create a CSV file and write data
        with open(filename, 'w', newline='') as csvfile:
            csv_writer = csv.DictWriter(csvfile, fieldnames=headers)
            csv_writer.writeheader()
            csv_writer.writerows(data)
        print(f"\033[0;32mData saved to {filename}")

    elif export_format == 'json':
        # Create a JSON file and write data
        with open(filename, 'w') as jsonfile:
            json.dump(data, jsonfile, indent=4)
        print(f"\033[0;32mData saved to {filename}")

    else:
        print(f"\033[0;31mInvalid export format. Please choose XLSX, CSV, or JSON.")

# Function to retrieve workspace metrics for a given location
def workspace_metrics(location_name, aggregation, export_format):
    location_id = get_location_id(location_name)

    if location_id:
        # Get a list of floor objects using get_floor_id function
        floor_list = get_floor_id(location_id)

        # Create a list to store the workspace metrics data
        workspace_metrics_data = []

        # Loop through metrics and workspaces to retrieve data and populate workspace_metrics_data list
        for floor in floor_list:
            floor_id = floor['id']
            floor_number = floor['floorNumber']

            # Get a list of workspace objects using get_workspace_id function
            workspace_list = get_workspace_id(location_id, floor_id)

            for workspace in workspace_list:
                workspace_id = workspace['id']
                workspace_name = get_workspace_name(workspace_id)  # Get the workspace name
                print(f"\033[0;38m{workspace_name} in progress...")

                for metric_name in METRIC_NAMES:
                    if metric_name == 'duration':
                        metrics_data = get_workspace_duration_metrics(workspace_id, aggregation, from_date_time_iso, to_date_time_iso)
                    else:
                        metrics_data = get_workspace_metrics(workspace_id, metric_name, aggregation, from_date_time_iso, to_date_time_iso)

                    if metrics_data:
                        # Define the capacity variable here
                        capacity = get_workspace_capacity(workspace_id)
                        
                        for metric in metrics_data.get('items', []):
                            start_time = metric.get('start', 'N/A')
                            end_time = metric.get('end', 'N/A')

                            if metric_name == 'duration':
                                duration = metric.get('duration', 'N/A')
                                workspace_metrics_data.append({
                                    'Workspace Name': workspace_name,
                                    'Floor Number': floor_number,
                                    'Capacity': capacity,
                                    'Metric Name': metric_name,
                                    'Start Date/Time': start_time,
                                    'End Date/Time': end_time,
                                    'Duration': duration,
                                    'Mean value': 'N/A',
                                    'Min value': 'N/A',
                                    'Max value': 'N/A'
                                })
                            else:
                                mean_value = metric.get('mean', 'N/A')
                                max_value = metric.get('max', 'N/A')
                                min_value = metric.get('min', 'N/A')
                                workspace_metrics_data.append({
                                    'Workspace Name': workspace_name,
                                    'Floor Number': floor_number,
                                    'Capacity': capacity,
                                    'Metric Name': metric_name,
                                    'Start Date/Time': start_time,
                                    'End Date/Time': end_time,
                                    'Duration': 'N/A',
                                    'Mean value': mean_value,
                                    'Min value': min_value,
                                    'Max value': max_value
                                })
                        print(f"\033[0;32m{metric_name} for {workspace_name} workspace done")
                print(f"\033[0;36m{workspace_name} workspace done \n")

        HEADERS_XLSX = ['Workspace Name', 'Floor Number', 'Capacity', 'Metric Name', 'Start Date/Time', 'End Date/Time', 'Duration', 'Mean value', 'Min value', 'Max value']

        # Define the filename for the export
        current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")  # Get the current date and time in the format 'YYYY-MM-DD_HH-MM-SS'
        filename = f'{current_datetime}_{location_name}_workspace_metrics_{aggregation}.{export_format}'

        # Export data based on the chosen format
        export_data(workspace_metrics_data, HEADERS_XLSX, filename, export_format)

    else:
        print(f"\033[0;31mUnable to get locationId.")

# Call the function to check the token
TOKEN = check_access_token(ACCESS_TOKEN, HEADERS)
print(TOKEN)

# Ask the user for the location name
LOCATION_NAME = input(f"\033[0;38mEnter the location name: ")
print('\n')

# Ask the user for aggregation choice
AGGREGATION_CHOICE = input("""Choose aggregation:
 1. hourly (the maximum time span is 48 hours)
 2. daily (the maximum time span is 30 days)
Enter your choice: """)
print('\n')

if AGGREGATION_CHOICE == '1':
    aggregation = 'hourly'
    max_time_span = HOURLY_MAX_TIME_SPAN
    to_date_time = datetime.now()
    to_date_time_iso = to_date_time.isoformat()
    from_date_time = to_date_time - timedelta(hours=max_time_span)
    from_date_time_iso = from_date_time.isoformat()
elif AGGREGATION_CHOICE == '2':
    aggregation = 'daily'
    max_time_span = DAILY_MAX_TIME_SPAN
    to_date_time = datetime.now()
    to_date_time_iso = to_date_time.isoformat()
    from_date_time = to_date_time - timedelta(days=max_time_span)
    from_date_time_iso = from_date_time.isoformat()
else:
    print(f"\033[0;31mInvalid aggregation choice. Please choose 1 for hourly or 2 for daily.")
    exit(1)

# Ask the user for the export format choice
EXPORT_FORMAT_CHOICE = input("""Choose export format:
 1. XLSX
 2. CSV
 3. JSON
Enter your choice: """).strip().lower()
print('\n')

if EXPORT_FORMAT_CHOICE == '1':
    export_format = 'xlsx'
elif EXPORT_FORMAT_CHOICE == '2':
    export_format = 'csv'
elif EXPORT_FORMAT_CHOICE == '3':
    export_format = 'json'
else:
    print(f"\033[0;31mInvalid export format choice. Please choose 1 for XLSX, 2 for CSV, or 3 for JSON.")
    exit(1)

# Retrieve workspace metrics
workspace_metrics(LOCATION_NAME, aggregation, export_format)