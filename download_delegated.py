import requests
import os
import json
from datetime import datetime

def get_access_token_delegated():
    """Get access token using refresh token (delegated permissions)"""
    tenant_id = os.environ['TENANT_ID']
    client_id = os.environ['CLIENT_ID']
    client_secret = os.environ['CLIENT_SECRET']
    refresh_token = os.environ['REFRESH_TOKEN']
    
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    data = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'client_secret': client_secret,
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/Files.Read.All offline_access'
    }
    
    response = requests.post(url, data=data)
    response.raise_for_status()
    
    token_data = response.json()
    
    # Update refresh token if a new one was issued
    if 'refresh_token' in token_data:
        print("ℹ️ New refresh token issued")
    
    return token_data['access_token']

def download_file_with_delegated_auth():
    """Download the file using delegated permissions"""
    access_token = get_access_token_delegated()
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }
    
    # Try multiple methods to find and download the file
    methods = [
        ("SharePoint Site Search", lambda: search_in_site(headers)),
        ("User Drive Search", lambda: search_in_user_drive(headers)),
        ("Direct Site Access", lambda: access_site_directly(headers))
    ]
    
    for method_name, method_func in methods:
        try:
            print(f"Trying {method_name}...")
            file_content = method_func()
            if file_content:
                print(f"✅ Success with {method_name}!")
                return file_content
        except Exception as e:
            print(f"❌ {method_name} failed: {e}")
    
    return None

def search_in_site(headers):
    """Search for the file in the SharePoint site"""
    # Search in the main SharePoint site
    search_url = "https://graph.microsoft.com/v1.0/sites/aphrcorg-my.sharepoint.com/drive/root/search(q='Active fellows PhD status')"
    
    response = requests.get(search_url, headers=headers)
    if response.status_code == 200:
        results = response.json()
        for item in results.get('value', []):
            if 'Active' in item.get('name', '') and item.get('name', '').endswith('.xlsx'):
                print(f"Found file: {item['name']}")
                return download_file_by_id(item['id'], headers, site_drive=True)
    
    return None

def search_in_user_drive(headers):
    """Search in nnjenga's personal drive"""
    search_url = "https://graph.microsoft.com/v1.0/users/nnjenga@aphrc.org/drive/root/search(q='Active fellows PhD status')"
    
    response = requests.get(search_url, headers=headers)
    if response.status_code == 200:
        results = response.json()
        for item in results.get('value', []):
            if 'Active' in item.get('name', '') and item.get('name', '').endswith('.xlsx'):
                print(f"Found file: {item['name']}")
                return download_file_by_id(item['id'], headers, user_drive="nnjenga@aphrc.org")
    
    return None

def access_site_directly(headers):
    """Try to access the file using known path"""
    # Try to construct the path based on the original URL
    file_path = "/personal/nnjenga_aphrc_org/Documents/CARTA Dashboard/ECRS/Active fellows PhD status.xlsx"
    encoded_path = requests.utils.quote(file_path, safe='/')
    
    url = f"https://graph.microsoft.com/v1.0/sites/aphrcorg-my.sharepoint.com/drive/root:{encoded_path}:/content"
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        print("Found file via direct path")
        return response.content
    
    return None

def download_file_by_id(file_id, headers, site_drive=False, user_drive=None):
    """Download file using its ID"""
    if user_drive:
        download_url = f"https://graph.microsoft.com/v1.0/users/{user_drive}/drive/items/{file_id}/content"
    elif site_drive:
        download_url = f"https://graph.microsoft.com/v1.0/sites/aphrcorg-my.sharepoint.com/drive/items/{file_id}/content"
    else:
        download_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
    
    response = requests.get(download_url, headers=headers)
    response.raise_for_status()
    return response.content

def main():
    try:
        print("🚀 Starting delegated authentication download...")
        
        file_content = download_file_with_delegated_auth()
        
        if file_content and len(file_content) > 1000:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"data/Active_fellows_PhD_status_{timestamp}.xlsx"
            
            os.makedirs("data", exist_ok=True)
            
            # Save timestamped version
            with open(filename, 'wb') as f:
                f.write(file_content)
            
            # Save latest version
            with open("data/Active_fellows_PhD_status_latest.xlsx", 'wb') as f:
                f.write(file_content)
            
            print(f"✅ Success! Downloaded {len(file_content):,} bytes")
            print(f"📁 Saved as: {filename}")
            print(f"📁 Latest: data/Active_fellows_PhD_status_latest.xlsx")
            
        else:
            print("❌ Download failed or file too small")
            exit(1)
            
    except Exception as e:
        print(f"💥 Error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
