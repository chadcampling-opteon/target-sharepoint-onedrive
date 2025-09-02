"""SharePoint OneDrive target class."""

from __future__ import annotations

from datetime import datetime
from typing import Any
from urllib.parse import urlparse

from singer_sdk import typing as th
from singer_sdk.target_base import Target

from target_sharepoint_onedrive.sinks import (
    SharePointOneDriveSink,
)


class TargetSharePointOneDrive(Target):
    """Target for SharePoint Document Libraries and OneDrive."""

    name = "target-sharepoint-onedrive"

    config_jsonschema = th.PropertiesList(
        # Authentication
        th.Property(
            "tenant_id",
            th.StringType(),
            required=False,
            title="Tenant ID",
            description="Azure AD tenant ID (optional if using environment-based authentication)",
        ),
        th.Property(
            "client_id",
            th.StringType(),
            required=False,
            title="Client ID",
            description="Azure AD application client ID (optional if using environment-based authentication)",
        ),
        th.Property(
            "client_secret",
            th.StringType(),
            required=False,
            secret=True,
            title="Client Secret",
            description="Azure AD application client secret (optional if using environment-based authentication)",
        ),
        th.Property(
            "destination_url",
            th.StringType(),
            required=True,
            title="Destination URL",
            description="SharePoint or OneDrive destination URL. For SharePoint: https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{site-name}:/ For OneDrive: https://graph.microsoft.com/v1.0/me/drive or https://graph.microsoft.com/v1.0/users/{user-id}/drive",
        ),
        th.Property(
            "folder_path",
            th.StringType(),
            title="Folder Path",
            description="Optional folder path within the document library or OneDrive",
            default="",
        ),
        # File Configuration
        th.Property(
            "file_naming_scheme",
            th.StringType(),
            title="File Naming Scheme",
            description="Scheme for naming output files. Supports variables: {stream_name}, {timestamp}, {batch_id}",
            default="{stream_name}_{timestamp}.csv",
        ),
        th.Property(
            "timestamp_format",
            th.StringType(),
            title="Timestamp Format",
            description="Format for timestamp in filename (strftime format)",
            default="%Y%m%d_%H%M%S",
        ),
        th.Property(
            "overwrite_files",
            th.BooleanType(),
            title="Overwrite Files",
            description="Whether to overwrite existing files with the same name",
            default=True,
        ),
        th.Property(
            "batch_size_rows",
            th.IntegerType(),
            title="Batch Size Rows",
            description="Maximum number of records per file",
            default=10000,
        ),
        # Stream Maps
        th.Property(
            "stream_maps",
            th.ObjectType(),
            title="Stream Maps",
            description="Config object for stream maps capability",
        ),
        th.Property(
            "stream_map_config",
            th.ObjectType(),
            title="Stream Map Config",
            description="User-defined config values to be used within map expressions",
        ),
        # Schema Flattening
        th.Property(
            "flattening_enabled",
            th.BooleanType(),
            title="Flattening Enabled",
            description="'True' to enable schema flattening and automatically expand nested properties",
            default=False,
        ),
        th.Property(
            "flattening_max_depth",
            th.IntegerType(),
            title="Flattening Max Depth",
            description="The max depth to flatten schemas",
            default=0,
        ),
    ).to_dict()

    default_sink_class = SharePointOneDriveSink

    def get_sink_class(self, stream_name: str) -> type[SharePointOneDriveSink]:
        """Get sink class for the stream.

        Args:
            stream_name: Name of the stream

        Returns:
            The sink class to use for this stream
        """
        return SharePointOneDriveSink

    def _validate_config(self, *, raise_errors: bool = True) -> list[str]:
        """Validate configuration input against the plugin configuration JSON schema.

        Args:
            raise_errors: Flag to throw an exception if any validation errors are found.

        Returns:
            A list of validation errors.

        Raises:
            ConfigValidationError: If raise_errors is True and validation fails.
        """
        errors = super()._validate_config(raise_errors=raise_errors)
        
        # Additional validation for destination URL
        config = self._config
        
        if config.get("destination_url"):
            url_errors = self._validate_destination_url(config["destination_url"])
            errors.extend(url_errors)
        
        if errors and raise_errors:
            from singer_sdk.exceptions import ConfigValidationError
            raise ConfigValidationError(
                f"Config validation failed: {'; '.join(errors)}",
                errors=errors
            )
        
        return errors

    def _validate_destination_url(self, url: str) -> list[str]:
        """Validate the destination URL format.

        Args:
            url: The destination URL to validate

        Returns:
            List of validation errors
        """
        errors = []
        
        try:
            parsed = urlparse(url)
            
            # Check if it's a valid Microsoft Graph URL
            if not url.startswith("https://graph.microsoft.com/v1.0/"):
                errors.append(
                    "Destination URL must start with 'https://graph.microsoft.com/v1.0/'"
                )
            
            # Check if it's a valid SharePoint or OneDrive URL
            is_sharepoint = url.startswith("https://graph.microsoft.com/v1.0/sites/")
            is_onedrive_me = url == "https://graph.microsoft.com/v1.0/me/drive"
            is_onedrive_user = url.startswith("https://graph.microsoft.com/v1.0/users/") and url.endswith("/drive")
            is_onedrive_direct = url.startswith("https://graph.microsoft.com/v1.0/drives/")
            
            if not any([is_sharepoint, is_onedrive_me, is_onedrive_user, is_onedrive_direct]):
                errors.append(
                    "Destination URL must be a valid SharePoint or OneDrive URL. "
                    "SharePoint: https://graph.microsoft.com/v1.0/sites/{domain}:/sites/{site-name}:/ "
                    "OneDrive: https://graph.microsoft.com/v1.0/me/drive, /users/{user-id}/drive, or /drives/{drive-id}"
                )
            
            # SharePoint-specific validation
            if is_sharepoint:
                if ":/sites/" not in url:
                    errors.append(
                        "SharePoint destination URL must contain ':/sites/' path segment"
                    )
                
                if not url.endswith(":/"):
                    errors.append(
                        "SharePoint destination URL must end with ':/'"
                    )

        except Exception as e:
            errors.append(f"Invalid URL format: {str(e)}")
        
        return errors

    def _detect_destination_type(self, url: str) -> str:
        """Detect destination type from URL.

        Args:
            url: The destination URL

        Returns:
            Destination type ('sharepoint' or 'onedrive')
        """
        if url.startswith("https://graph.microsoft.com/v1.0/sites/"):
            return "sharepoint"
        else:
            return "onedrive"



if __name__ == "__main__":
    TargetSharePointOneDrive.cli()
