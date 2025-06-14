import requests
import os
from datetime import datetime
import msal

def get_access_token():
    """Get access token for Microsoft Graph API"""
    tenant_id = os.environ['TENANT_ID']
    client_id = os.environ['CLIENT_ID']
    client_secret = os.environ['CLIENT_SECRET']
    
    # Create a confidential client application
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret
    )
    
    # Get token for Microsoft Graph
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Failed to get access token: {result.get('error_description', 'Unknown error')}")

def download_file():
    """Download the Excel file from SharePoint"""
    source_url = os.environ['SOURCE_URL']
    
    try:
        response = requests.get(source_url, timeout=30)
        response.raise_for_status()
        return response.content
    except requests.RequestException as e:
        print(f"Error downloading file: {e}")
        raise

def upload_to_onedrive(file_content, access_token):
    """Upload file to OneDrive"""
    # Create filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Active_fellows_PhD_status_{timestamp}.xlsx"
    
    # Microsoft Graph API endpoint for OneDrive
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/CARTA_Dashboard/{filename}:/content"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    
    response = requests.put(upload_url, headers=headers, data=file_content)
    
    if response.status_code == 201:
        print(f"File uploaded successfully: {filename}")
        return response.json()
    else:
        print(f"Upload failed: {response.status_code} - {response.text}")
        raise Exception(f"Upload failed: {response.status_code}")

def main():
    try:
        print("Starting download and upload process...")
        
        # Get access token
        print("Getting access token...")
        access_token = get_access_token()
        
        # Download file
        print("Downloading file...")
        file_content = download_file()
        print(f"Downloaded {len(file_content)} bytes")
        
        # Upload to OneDrive
        print("Uploading to OneDrive...")
        result = upload_to_onedrive(file_content, access_token)
        
        print("Process completed successfully!")
        print(f"File URL: {result.get('webUrl', 'N/A')}")
        
    except Exception as e:
        print(f"Error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
