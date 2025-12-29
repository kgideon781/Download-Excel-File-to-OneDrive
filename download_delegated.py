import requests
import os
import json
import glob
from datetime import datetime

def get_access_token_delegated():
    """Get access token using refresh token (delegated permissions) with improved error handling"""
    tenant_id = os.environ['TENANT_ID'].strip()
    client_id = os.environ['CLIENT_ID'].strip()
    client_secret = os.environ['CLIENT_SECRET'].strip()
    refresh_token = os.environ['REFRESH_TOKEN'].strip()
    
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    data = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'client_secret': client_secret,
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/Files.Read.All offline_access'
    }
    
    print("[*] Requesting access token...")
    print(f"[*] Token endpoint: {url}")
    print(f"[*] Client ID: {client_id[:8]}...{client_id[-4:]}")
    print(f"[*] Refresh token length: {len(refresh_token)} chars")
    
    try:
        response = requests.post(url, data=data, timeout=30)
        
        if response.status_code != 200:
            print(f"[X] Authentication failed: {response.status_code}")
            try:
                error_data = response.json()
                print(f"[X] Error: {error_data.get('error', 'Unknown')}")
                print(f"[X] Description: {error_data.get('error_description', 'No description')}")
            except:
                print(f"[X] Response: {response.text}")
            
            response.raise_for_status()
        
        token_data = response.json()
        
        if 'refresh_token' in token_data:
            new_refresh = token_data['refresh_token']
            print("[!] New refresh token issued!")
            print(f"[!] UPDATE GitHub Secret REFRESH_TOKEN with:\n{new_refresh}\n")
            
            # Save to file for convenience
            with open('NEW_REFRESH_TOKEN.txt', 'w', encoding='utf-8') as f:
                f.write(f"New refresh token issued: {datetime.now()}\n\n")
                f.write(f"{new_refresh}\n\n")
                f.write("ACTION REQUIRED:\n")
                f.write("Update GitHub Secret 'REFRESH_TOKEN' with the value above\n")
            print("[*] New token also saved to: NEW_REFRESH_TOKEN.txt")
        
        print("[+] Access token obtained successfully")
        return token_data['access_token']
        
    except requests.exceptions.HTTPError as e:
        print(f"\n[X] HTTP Error: {e}")
        raise
    except Exception as e:
        print(f"\n[X] Unexpected error: {e}")
        raise

# 📋 CONFIGURATION: File management settings
FILE_MANAGEMENT = {
    'keep_versions': 2,  # Keep latest + 1 previous version
    'archive_old_files': True,  # Move old files to archive folder
    'create_changelog': True  # Track what changed
}

FILES_TO_DOWNLOAD = [
    {
        'search_terms': ['Active fellows PhD status'],
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
    {
        'search_terms': ['Postdoctoral awards to CARTA graduates'],
        'filename_contains': 'Postdoctoral awards to CARTA graduates',
        'output_name': 'Postdocs',
        'description': 'Postdoctoral awards to CARTA graduates DB'
    },
    {
        'search_terms': ['Fellows secured extra grants'],
        'filename_contains': 'Fellows secured extra grants',
        'output_name': 'Extra Grants',
        'description': 'Fellows secured extra grants DB'
    }
]

def search_for_file(headers, search_config):
    """Search for a specific file using multiple strategies"""
    search_terms = search_config['search_terms']
    filename_contains = search_config['filename_contains']
    
    print(f"[*] Searching for: {search_config['description']}")
    
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
                        print(f"    [+] Found: {item['name']}")
                        return download_file_by_id(item['id'], headers, user_drive="nnjenga@aphrc.org")
        except Exception as e:
            print(f"    [X] Search failed for '{term}': {e}")
    
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
                        print(f"    [+] Found in SharePoint: {item['name']}")
                        return download_file_by_id(item['id'], headers, site_drive=True)
        except Exception as e:
            print(f"    [X] SharePoint search failed for '{term}': {e}")
    
    print(f"    [X] File not found matching: {filename_contains}")
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

def manage_file_versions(output_name):
    """Manage file versions - keep only the specified number of versions"""
    data_dir = "data"
    archive_dir = "data/archive"
    
    # Create directories if they don't exist
    os.makedirs(data_dir, exist_ok=True)
    if FILE_MANAGEMENT['archive_old_files']:
        os.makedirs(archive_dir, exist_ok=True)
    
    # Find all versions of this file
    pattern = f"{data_dir}/{output_name}_*.xlsx"
    existing_files = glob.glob(pattern)
    
    # Sort by modification time (newest first)
    existing_files.sort(key=os.path.getmtime, reverse=True)
    
    keep_count = FILE_MANAGEMENT['keep_versions']
    
    if len(existing_files) >= keep_count:
        files_to_remove = existing_files[keep_count:]
        
        for old_file in files_to_remove:
            if FILE_MANAGEMENT['archive_old_files']:
                # Move to archive
                archive_path = os.path.join(archive_dir, os.path.basename(old_file))
                os.rename(old_file, archive_path)
                print(f"    [*] Archived: {os.path.basename(old_file)}")
            else:
                # Delete the file
                os.remove(old_file)
                print(f"    [*] Deleted old version: {os.path.basename(old_file)}")

def check_if_file_changed(new_content, output_name):
    """Check if the new file is different from the current latest version"""
    latest_file = f"data/{output_name}_latest.xlsx"
    
    if not os.path.exists(latest_file):
        return True  # No previous file, so it's "changed"
    
    with open(latest_file, 'rb') as f:
        old_content = f.read()
    
    # Simple comparison - files are different if sizes differ or content differs
    if len(new_content) != len(old_content):
        return True
    
    return new_content != old_content

def save_file_with_version_control(file_content, output_name, description):
    """Save file with smart version control"""
    if not file_content or len(file_content) < 1000:
        print(f"    [X] File too small or empty: {description}")
        return False
    
    # Check if file actually changed
    if not check_if_file_changed(file_content, output_name):
        print(f"    [*] No changes detected for: {description}")
        return True
    
    print(f"    [+] Changes detected - updating: {description}")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create data directory
    os.makedirs("data", exist_ok=True)
    
    # Clean up old versions first
    manage_file_versions(output_name)
    
    # Save new timestamped version
    timestamped_filename = f"data/{output_name}_{timestamp}.xlsx"
    with open(timestamped_filename, 'wb') as f:
        f.write(file_content)
    
    # Update latest version
    latest_filename = f"data/{output_name}_latest.xlsx"
    with open(latest_filename, 'wb') as f:
        f.write(file_content)
    
    print(f"    [+] Saved {len(file_content):,} bytes")
    print(f"    [*] Current: {latest_filename}")
    print(f"    [*] Backup: {timestamped_filename}")
    
    # Update changelog
    if FILE_MANAGEMENT['create_changelog']:
        update_changelog(output_name, description, timestamp)
    
    return True

def update_changelog(output_name, description, timestamp):
    """Update the changelog with what was downloaded"""
    changelog_file = "data/CHANGELOG.md"
    
    # Read existing changelog
    changelog_content = ""
    if os.path.exists(changelog_file):
        with open(changelog_file, 'r', encoding='utf-8') as f:
            changelog_content = f.read()
    
    # Prepare new entry
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S UTC")
    new_entry = f"\n## {timestamp}\n**Date:** {date_str}  \n**File:** {description} (`{output_name}_latest.xlsx`)  \n**Status:** Updated\n"
    
    # Add to changelog
    if "# File Download Changelog" not in changelog_content:
        changelog_content = "# File Download Changelog\n\nThis file tracks all file downloads and updates.\n" + new_entry
    else:
        # Insert after the header
        parts = changelog_content.split('\n', 3)
        if len(parts) >= 3:
            changelog_content = '\n'.join(parts[:3]) + new_entry + '\n'.join(parts[3:])
        else:
            changelog_content += new_entry
    
    # Write back
    with open(changelog_file, 'w', encoding='utf-8') as f:
        f.write(changelog_content)

def download_all_files():
    """Download all configured files with version control"""
    print("\n" + "=" * 70)
    print("CARTA DASHBOARD - FILE DOWNLOAD")
    print("=" * 70)
    
    access_token = get_access_token_delegated()
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json'
    }
    
    successful_downloads = 0
    failed_downloads = 0
    unchanged_files = 0
    
    print(f"\n[*] Starting multi-file download with version control...")
    print(f"[*] Keeping {FILE_MANAGEMENT['keep_versions']} versions per file")
    print("=" * 70)
    
    for i, file_config in enumerate(FILES_TO_DOWNLOAD, 1):
        print(f"\n[{i}/{len(FILES_TO_DOWNLOAD)}] {file_config['description']}")
        print("-" * 40)
        
        try:
            file_content = search_for_file(headers, file_config)
            
            if file_content:
                if save_file_with_version_control(file_content, file_config['output_name'], file_config['description']):
                    successful_downloads += 1
                else:
                    failed_downloads += 1
            else:
                print(f"    [X] Could not download: {file_config['description']}")
                failed_downloads += 1
                
        except Exception as e:
            print(f"    [X] Error downloading {file_config['description']}: {e}")
            failed_downloads += 1
    
    print("\n" + "=" * 70)
    print("DOWNLOAD SUMMARY")
    print("=" * 70)
    print(f"[+] Successful: {successful_downloads}")
    print(f"[X] Failed: {failed_downloads}")
    print(f"[*] Total files processed: {len(FILES_TO_DOWNLOAD)}")
    print(f"\n[*] File organization:")
    print(f"    - Latest versions: data/*_latest.xlsx")
    print(f"    - Backup versions: data/*_YYYYMMDD_HHMMSS.xlsx")
    if FILE_MANAGEMENT['archive_old_files']:
        print(f"    - Archived files: data/archive/")
    if FILE_MANAGEMENT['create_changelog']:
        print(f"    - Change history: data/CHANGELOG.md")
    
    return successful_downloads, failed_downloads

def main():
    try:
        successful, failed = download_all_files()
        
        if failed > 0:
            print(f"\n[!] Some downloads failed. Check the logs above.")
            exit(1)
        else:
            print(f"\n[+] All files processed successfully!")
            
    except Exception as e:
        print(f"\n[X] Critical error: {e}")
        import traceback
        traceback.print_exc()
        exit(1)

if __name__ == "__main__":
    main()
