import os
import requests
from flask import Flask, request, jsonify
from msal import ConfidentialClientApplication

"""
Flask application for managing files in SharePoint using Microsoft Graph API.
This application provides endpoints for uploading and downloading files from SharePoint,
with automatic folder creation and path management.

Required Environment Variables:
    - SP_CLIENT_ID: SharePoint application client ID
    - SP_CLIENT_SECRET: SharePoint application client secret
    - SP_TENANT_ID: SharePoint tenant ID
    - SP_SITE_ID: SharePoint site ID
"""

# SharePoint Configuration
SP_CLIENT_ID = os.getenv("SP_CLIENT_ID", "")
SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "")
SP_TENANT_ID = os.getenv("SP_TENANT_ID", "")
SP_SITE_ID = os.getenv("SP_SITE_ID", "")

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'  # Change this in production
app.config['DEBUG'] = True

# Authentication Functions
def fetch_header():
    """
    Get authentication header for Microsoft Graph API requests.
    
    Returns:
        dict: Header containing bearer token for authentication
    """
    authority_url = f"https://login.microsoftonline.com/{SP_TENANT_ID}"
    client_app = ConfidentialClientApplication(
        SP_CLIENT_ID,
        authority=authority_url,
        client_credential=SP_CLIENT_SECRET,
    )
    token_response = client_app.acquire_token_for_client(
        scopes=['https://graph.microsoft.com/.default']
    )
    return {'Authorization': f'Bearer {token_response["access_token"]}'}

# SharePoint Operations
def create_folder_by_path(path):
    """
    Create folder structure in SharePoint recursively.
    
    Args:
        path (str): Path where folders should be created (e.g., 'folder1/folder2/folder3')
    
    Returns:
        str: URL of the created folder structure, or -1 if creation fails
    """
    header = fetch_header()
    parts = path.split('/')
    current_path = "drive/root:"

    for part in parts:
        if not part:  # Skip empty parts
            continue
            
        # Check if folder exists
        check_url = f"https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/{current_path}/{part}"
        check_response = requests.get(check_url, headers=header)
        
        if check_response.status_code in [200, 201]:
            print(f"Folder '{part}' exists.")
            current_path += f"/{part}"
        else:
            # Create new folder
            create_url = f"https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/{current_path}:/children"
            header['Content-Type'] = 'application/json'
            folder_data = {
                "name": part,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename"
            }
            create_response = requests.post(create_url, headers=header, json=folder_data)
            
            if create_response.status_code in [200, 201]:
                print(f"Folder '{part}' created.")
                current_path += f"/{part}"
            else:
                print(f"Error creating folder '{part}':", create_response.json())
                return -1
                
    return create_url.replace(":/children", f"/{part}:/children")

def fetch_file(path, filename):
    """
    Fetch file download URL from SharePoint.
    
    Args:
        path (str): Path to the file in SharePoint
        filename (str): Name of the file to fetch
    
    Returns:
        str: Download URL for the file, or -1 if fetch fails
    """
    try:
        header = fetch_header()
        path = clean_sharepoint_path(path)
        file_path = f"https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/drive/root:/{path}/{filename}"
        
        response = requests.get(file_path, headers=header)
        
        # Handle different response status codes
        if response.status_code == 404:
            print(f"File not found: {path}/{filename}")
            return -1
        elif response.status_code == 401:
            print("Authentication failed. Token might be expired.")
            return -1
        elif response.status_code == 403:
            print("Permission denied to access file.")
            return -1
        elif response.status_code not in [200, 201]:
            print(f"Unexpected error {response.status_code}: {response.text}")
            return -1

        data = response.json()
        download_url = data.get('@microsoft.graph.downloadUrl')
        if not download_url:
            print("Download URL not found in response")
            return -1
            
        return download_url

    except Exception as e:
        print(f"Error fetching file: {str(e)}")
        return -1

def upload_file(path, filename, file_contents):
    """
    Upload file to SharePoint.
    
    Args:
        path (str): Path where file should be uploaded
        filename (str): Name of the file to upload
        file_contents (bytes): Contents of the file
    
    Returns:
        str: Path of uploaded file, or -1 if upload fails
    """
    header = fetch_header()
    root_path = f"https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/drive/root:"
    file_path = f"{root_path}/{path}/{filename}:/content"
    header['Content-Type'] = 'text/plain'
    
    response = requests.put(file_path, headers=header, data=file_contents)
    if response.status_code in [200, 201]:
        print(f"Successfully uploaded {filename}")
        return file_path
        
    print(f"Failed to upload {filename}")
    return -1

def clean_sharepoint_path(path):
    """
    Clean and normalize SharePoint path.
    
    Args:
        path (str): Raw path to clean
    
    Returns:
        str: Cleaned path without extra slashes or content suffix
    """
    path = path.replace(':/content', '')
    path = path.strip('/')
    return '/'.join(filter(None, path.split('/')))

# Flask Routes
@app.route('/upload', methods=['POST'])
def upload_file_endpoint():
    """
    Endpoint for file uploads to SharePoint.
    
    Expected form data:
        - file: File to upload
        - path: (optional) Path where file should be uploaded
    
    Returns:
        JSON response with upload status and path
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
            
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Empty filename'}), 400
            
        path = clean_sharepoint_path(request.form.get('path', ''))
        
        # Ensure folder structure exists
        folder_path = create_folder_by_path(path)
        if folder_path == -1:
            return jsonify({'error': 'Failed to create folder structure'}), 500
            
        # Upload file
        result = upload_file(path, file.filename, file.read())
        if result == -1:
            return jsonify({'error': 'Failed to upload file'}), 500
            
        return jsonify({
            'message': 'File uploaded successfully',
            'path': result
        }), 201
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/fetch/<path:filepath>')
def fetch_file_endpoint(filepath):
    """
    Endpoint to fetch file download URL from SharePoint.
    
    Args:
        filepath: Full path to file including filename
    
    Returns:
        JSON response with download URL and file details
    """
    try:
        # Split path into directory path and filename
        path_parts = filepath.rsplit('/', 1)
        filename = path_parts[-1]
        path = path_parts[0] if len(path_parts) > 1 else ''
        
        path = clean_sharepoint_path(path)
        download_url = fetch_file(path, filename)
        
        if download_url == -1:
            return jsonify({'error': 'Failed to fetch file'}), 404
            
        return jsonify({
            'download_url': download_url,
            'filename': filename,
            'path': path
        }), 200
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Run the application
if __name__ == '__main__':
    app.run()