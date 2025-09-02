"""SharePoint and OneDrive client for Microsoft Graph API operations."""

from __future__ import annotations

import logging
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd
import requests
from azure.identity import ClientSecretCredential, DefaultAzureCredential

logger = logging.getLogger(__name__)


class MSGraphClient:
    """Client for interacting with SharePoint document libraries and OneDrive via Microsoft Graph API."""

    def __init__(
        self,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        destination_url: str = "",
        folder_path: str = "",
    ) -> None:
        """Initialize the client.

        Args:
            tenant_id: Azure AD tenant ID (optional if using environment-based authentication)
            client_id: Azure AD application client ID (optional if using environment-based authentication)
            client_secret: Azure AD application client secret (optional if using environment-based authentication)
            destination_url: SharePoint or OneDrive destination URL
            folder_path: Optional folder path within the document library or OneDrive
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.destination_url = destination_url
        self.folder_path = folder_path.strip("/")
        
        # Initialize credentials
        if tenant_id and client_id and client_secret:
            # Use explicit credentials if provided
            self.credential = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
            )
            logger.info("Using explicit Azure credentials from config")
        else:
            # Use default credential chain (environment variables, managed identity, etc.)
            self.credential = DefaultAzureCredential()
            logger.info("Using Azure default credential chain (environment variables, managed identity, etc.)")
        
        # Create requests session with authentication
        self.session = requests.Session()
        self.session.auth = self._get_auth_handler()
        
        # Detect destination type and resolve drive information from URL
        self.destination_type = self._detect_destination_type(destination_url)
        self._resolve_drive_from_url()
    
    def _get_auth_handler(self):
        """Get authentication handler for requests."""
        class AzureAuthHandler:
            def __init__(self, credential):
                self.credential = credential
            
            def __call__(self, request):
                # Get token for Microsoft Graph scope
                token = self.credential.get_token("https://graph.microsoft.com/.default")
                request.headers["Authorization"] = f"Bearer {token.token}"
                return request
        
        return AzureAuthHandler(self.credential)
        
    def _detect_destination_type(self, url: str) -> str:
        return "sharepoint" if url.startswith("https://graph.microsoft.com/v1.0/sites/") else "onedrive"
        
    def _resolve_drive_from_url(self) -> None:
        """Resolve drive information from destination URL."""
        if not self.destination_url:
            raise ValueError("destination_url is required")
        
        try:
            if self.destination_type == "sharepoint":
                self._resolve_sharepoint_drive()
            else:
                self._resolve_onedrive_drive()
                
        except Exception as e:
            logger.error(f"Failed to resolve drive from URL: {str(e)}")
            raise
    
    def _resolve_sharepoint_drive(self) -> None:
        """Resolve SharePoint site and drive IDs from destination URL."""
        # Extract site identifier from the destination URL
        # Example: https://graph.microsoft.com/v1.0/sites/myorg.sharepoint.com:/sites/MySite:/
        site_identifier = self.destination_url.replace("https://graph.microsoft.com/v1.0/sites/", "")
        
        # Get the default drive directly
        logger.info(f"Resolving SharePoint drive for site: {site_identifier}")
        
        try:
            # Get the default drive
            drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_identifier}/drive"
            drive_response = self.session.get(drive_url)
            drive_response.raise_for_status()
            drive_data = drive_response.json()
            
            self.site_id = site_identifier
            self.drive_id = drive_data['id']
            
            logger.info(f"Resolved SharePoint site_id: {self.site_id}, drive_id: {self.drive_id}")
            
        except Exception as e:
            logger.error(f"Failed to resolve SharePoint drive: {str(e)}")
            raise
    
    def _resolve_onedrive_drive(self) -> None:
        """Resolve OneDrive drive information from destination URL."""
        if self.destination_url == "https://graph.microsoft.com/v1.0/me/drive":
            # Current user's OneDrive
            logger.info("Resolving current user's OneDrive")
            
            try:
                drive_response = self.session.get(self.destination_url)
                drive_response.raise_for_status()
                drive_data = drive_response.json()
                
                self.drive_id = drive_data['id']
                self.drive_type = "me"
                
            except Exception as e:
                logger.error(f"Failed to resolve current user's OneDrive: {str(e)}")
                raise
            
        elif self.destination_url.startswith("https://graph.microsoft.com/v1.0/users/"):
            # Specific user's OneDrive
            user_id = self.destination_url.split("/users/")[1].split("/drive")[0]
            logger.info(f"Resolving OneDrive for user: {user_id}")
            
            try:
                drive_response = self.session.get(self.destination_url)

                drive_response.raise_for_status()
                drive_data = drive_response.json()
                
                self.drive_id = drive_data['id']
                self.drive_type = "user"
                self.user_id = user_id
                
            except Exception as e:
                logger.error(f"Failed to resolve OneDrive for user {user_id}: {str(e)}")
                raise
            
        elif self.destination_url.startswith("https://graph.microsoft.com/v1.0/drives/"):
            # Direct drive ID
            drive_id = self.destination_url.split("/drives/")[1]
            logger.info(f"Using direct drive ID: {drive_id}")
            self.drive_id = drive_id
            self.drive_type = "direct"
            
        else:
            raise ValueError(f"Invalid OneDrive destination URL format: {self.destination_url}")
        
        logger.info(f"Resolved OneDrive drive_id: {self.drive_id}")
        
    
    def upload_csv_file(
        self,
        data: list[dict],
        filename: str,
        overwrite: bool = True,
    ) -> str:
        """Upload a CSV file to SharePoint or OneDrive.

        Args:
            data: List of dictionaries representing the data to write
            filename: Name of the file to create
            overwrite: Whether to overwrite existing files

        Returns:
            The URL of the uploaded file

        Raises:
            Exception: If upload fails
        """
        try:
            logger.debug(f"Starting CSV upload: filename={filename}, data_count={len(data)}, destination_type={self.destination_type}")
            
            # Validate data is not empty
            if not data:
                raise ValueError("No data to upload")
            
            # Convert data to DataFrame and then to CSV
            df = pd.DataFrame(data)
            logger.debug(f"Created DataFrame with shape: {df.shape}")
            
            # Create temporary file
            with tempfile.NamedTemporaryFile(
                mode="w",
                suffix=".csv",
                delete=False,
                encoding="utf-8",
            ) as temp_file:
                df.to_csv(temp_file.name, index=False)
                temp_file_path = temp_file.name
                logger.debug(f"Created temporary CSV file: {temp_file_path}")
            
            # Upload the file
            file_url = self._upload_file(
                local_file_path=temp_file_path,
                remote_filename=filename,
                overwrite=overwrite,
            )
            
            # Clean up temporary file
            Path(temp_file_path).unlink()
            logger.debug(f"Cleaned up temporary file: {temp_file_path}")
            
            destination_name = "OneDrive" if self.destination_type == "onedrive" else "SharePoint"
            logger.info(f"Successfully uploaded {filename} to {destination_name}")
            return file_url
            
        except Exception as e:
            logger.error(f"Failed to upload {filename}: {str(e)}")
            raise
    
    def _upload_file(
        self,
        local_file_path: str,
        remote_filename: str,
        overwrite: bool = True,
    ) -> str:
        """Upload a file using Microsoft Graph API.

        Args:
            local_file_path: Path to the local file
            remote_filename: Name for the file
            overwrite: Whether to overwrite existing files

        Returns:
            The URL of the uploaded file
        """
        # Construct the full path
        if self.folder_path:
            full_path = f"{self.folder_path}/{remote_filename}"
        else:
            full_path = remote_filename
        
        logger.debug(f"Uploading file: local_path={local_file_path}, remote_path={full_path}, overwrite={overwrite}")
        
        # Ensure folder exists before uploading
        if self.folder_path:
            logger.debug(f"Ensuring folder exists: {self.folder_path}")
            self._ensure_folder_exists(self.folder_path)
        
        # For now, use simple upload for all files
        # In production, you might want to implement chunked upload for large files
        return self._upload_simple_file(local_file_path, full_path, overwrite)
    
    
    
    def _upload_simple_file(
        self,
        local_file_path: str,
        remote_path: str,
        overwrite: bool,
    ) -> str:
        """Upload a file using simple upload.

        Args:
            local_file_path: Path to the local file
            remote_path: Path in destination
            overwrite: Whether to overwrite existing files

        Returns:
            The URL of the uploaded file
        """
        # Read file content
        with open(local_file_path, "rb") as file:
            file_content = file.read()
        
        try:
            logger.debug(f"Uploading file: drive_id={self.drive_id}, path={remote_path}")
            
            # Construct the URL using direct drive access (works for both SharePoint and OneDrive)
            upload_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{remote_path}:/content"
            if overwrite:
                upload_url += "?@microsoft.graph.conflictBehavior=replace"
            
            # Upload the file
            headers = {"Content-Type": "application/octet-stream"}
            response = self.session.put(upload_url, data=file_content, headers=headers)
            response.raise_for_status()
            
            drive_item = response.json()
            return drive_item.get('webUrl', f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{remote_path}")
            
        except Exception as e:
            logger.error(f"Failed to upload file {remote_path}: {str(e)}")
            raise
    
    def _ensure_folder_exists(self, folder_path: str) -> None:
        """Ensure that the specified folder exists, creating it if necessary.

        Args:
            folder_path: The folder path to ensure exists
        """
        if not folder_path:
            return
        
        try:
            logger.debug(f"Ensuring folder exists: drive_id={self.drive_id}, path={folder_path}")
            
            try:
                # Check if folder exists using direct drive access
                check_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{folder_path}"
                response = self.session.get(check_url)
                response.raise_for_status()
                logger.debug(f"Folder {folder_path} already exists")
                
            except Exception as e:
                # Folder doesn't exist, create it
                logger.debug(f"Creating folder {folder_path}: {str(e)}")
                
                create_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/children"
                folder_name = folder_path.split("/")[-1]
                
                folder_data = {
                    "name": folder_name,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "rename"
                }
                
                headers = {"Content-Type": "application/json"}
                response = self.session.post(create_url, json=folder_data, headers=headers)
                response.raise_for_status()
                logger.debug(f"Successfully created folder {folder_path}")
            
        except Exception as e:
            logger.warning(f"Could not ensure folder exists {folder_path}: {str(e)}")
            # Continue anyway, the upload might still work 