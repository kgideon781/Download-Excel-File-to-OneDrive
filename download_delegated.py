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

# 📋 CONFIGURATION: Add your files here
FILES_TO_DOWNLOAD = [
    {
        'search_terms': ['Active fellows PhD  status'],
        'filename_contains': 'Active',
        'output_name': 'Active_fellows_PhD_status',
        'description': 'PhD Fellows Status Report'
    },
    {
        'search_terms': ['Cohort 1-10 Demographics'],
        'filename_contains': 'Cohort 1-10',
        'output_name': 'Cohort_1_10_Demographics',
        'description': 'Cohort 1-10 Demographics'
    },
    {
        'search_terms': ['institutional participant list'],
        'filename_contains': 'institutional participant list',
        'output_name': 'Institutionalization',
        'description': 'Institutional Achievements DB'
    },
    # Add more files here following the same pattern
    # {
    #     'search_terms': ['your', 'search', 'terms'],
    #     'filename_contains': 'part_of_filename',
    #     'output_name': 'Output_File_Name',
    #     'description': 'Human readable description'
    # }
]

def search_for_file(headers, search_config):
    """Search for a specific file using multiple strategies"""
    search_terms = search_config['search_terms']
    filename_contains = search_config['filename_contains']
    
    print(f"🔍 Searching for: {search_config['description']}")
    
    # Strategy 1: Search in nnjenga's personal drive
    for term in search_terms:
        search_url = f"https://graph.microsoft.com/v1.0/users/nnjenga@aphrc.org/drive/root/search(q='{term}')"
        
        try:
            response = requests.get(search_url, headers=headers)
            if response.status_code == 200:
                results = response.json()
                for item in results.get('value', []):
                    filename = item.get('name', '').lower()
                    if (filename_contains.lower() in filename and 
                        (filename.endswith('.xlsx') or filename.endswith('.xls'))):
                        print(f"   ✅ Found: {item['name']}")
                        return download_file_by_id(item['id'], headers, user_drive="nnjenga@aphrc.org")
        except Exception as e:
            print(f"   ❌ Search failed for '{term}': {e}")
    
    # Strategy 2: Search in SharePoint site
    for term in search_terms:
        search_url = f"https://graph.microsoft.com/v1.0/sites/aphrcorg-my.sharepoint.com/drive/root/search(q='{term}')"
        
        try:
            response = requests.get(search_url, headers=headers)
            if response.status_code == 200:
                results = response.json()
                for item in results.get('value', []):
                    filename = item.get('name', '').lower()
                    if (filename_contains.lower() in filename and 
                        (filename.endswith('.xlsx') or filename.endswith('.xls'))):
                        print(f"   ✅ Found in SharePoint: {item['name']}")
                        return download_file_by_id(item['id'], headers, site_drive=True)
        except Exception as e:
            print(f"   ❌ SharePoint search failed for '{term}': {e}")
    
    print(f"   ❌ File not found matching: {filename_contains}")
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

def save_file(file_content, output_name, description):
    """Save file with timestamp and latest versions"""
    if not file_content or len(file_content) < 1000:
        print(f"   ❌ File too small or empty: {description}")
        return False
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create data directory
    os.makedirs("data", exist_ok=True)
    
    # Save timestamped version
    timestamped_filename = f"data/{output_name}_{timestamp}.xlsx"
    with open(timestamped_filename, 'wb') as f:
        f.write(file_content)
    
    # Save latest version
    latest_filename = f"data/{output_name}_latest.xlsx"
    with open(latest_filename, 'wb') as f:
        f.write(file_content)
    
    print(f"   ✅ Saved {len(file_content):,} bytes")
    print(f"   📁 Timestamped: {timestamped_filename}")
    print(f"   📁 Latest: {latest_filename}")
    return True

def download_all_files():
    """Download all configured files"""
    access_token = get_access_token_delegated()
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }
    
    successful_downloads = 0
    failed_downloads = 0
    
    print("🚀 Starting multi-file download...")
    print("=" * 60)
    
    for i, file_config in enumerate(FILES_TO_DOWNLOAD, 1):
        print(f"\n📁 File {i}/{len(FILES_TO_DOWNLOAD)}: {file_config['description']}")
        print("-" * 40)
        
        try:
            file_content = search_for_file(headers, file_config)
            
            if file_content:
                if save_file(file_content, file_config['output_name'], file_config['description']):
                    successful_downloads += 1
                else:
                    failed_downloads += 1
            else:
                print(f"   ❌ Could not download: {file_config['description']}")
                failed_downloads += 1
                
        except Exception as e:
            print(f"   💥 Error downloading {file_config['description']}: {e}")
            failed_downloads += 1
    
    print("\n" + "=" * 60)
    print("📊 DOWNLOAD SUMMARY")
    print("=" * 60)
    print(f"✅ Successful: {successful_downloads}")
    print(f"❌ Failed: {failed_downloads}")
    print(f"📁 Total files processed: {len(FILES_TO_DOWNLOAD)}")
    
    return successful_downloads, failed_downloads

def main():
    try:
        successful, failed = download_all_files()
        
        if failed > 0:
            print(f"\n⚠️ Some downloads failed. Check the logs above.")
            exit(1)
        else:
            print(f"\n🎉 All files downloaded successfully!")
            
    except Exception as e:
        print(f"💥 Critical error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
