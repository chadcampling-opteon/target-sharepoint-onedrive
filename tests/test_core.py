"""Tests standard target features using the built-in SDK tests library."""

from __future__ import annotations

import typing as t
from unittest.mock import patch

import pytest
from singer_sdk.testing import get_target_test_class

from target_sharepoint_onedrive.target import TargetSharePointOneDrive

# Initialize minimal target config for testing
SAMPLE_CONFIG: dict[str, t.Any] = {
    "destination_url": "https://graph.microsoft.com/v1.0/me/drive",
    "folder_path": "test_data",
    "file_naming_scheme": "{stream_name}_{timestamp}.csv",
    "timestamp_format": "%Y%m%d_%H%M%S",
    "overwrite_files": True,
    "batch_size_rows": 1000,
    "flattening_enabled": False,
    "flattening_max_depth": 0,
}


# Run standard built-in target tests from the SDK:
StandardTargetTests = get_target_test_class(
    target_class=TargetSharePointOneDrive,
    config=SAMPLE_CONFIG,
)


class TestTargetSharePointOneDrive(StandardTargetTests):  # type: ignore[misc, valid-type]
    """Standard Target Tests."""

    @pytest.fixture(scope="class")
    def resource(self):  # noqa: ANN201
        """Generic external resource.

        This fixture is useful for setup and teardown of external resources,
        such output folders, tables, buckets etc. for use during testing.

        Example usage can be found in the SDK samples test suite:
        https://github.com/meltano/sdk/tree/main/tests/packages
        """
        return "resource"

    def setup_method(self):
        """Set up test method with mocked SharePoint client."""
        # Mock the SharePoint client initialization to prevent real HTTP requests
        self.sharepoint_patcher = patch('target_sharepoint_onedrive.ms_graph_client.MSGraphClient._resolve_drive_from_url')
        self.mock_resolve_drive = self.sharepoint_patcher.start()
        
        # Also mock the upload method to prevent real uploads
        self.upload_patcher = patch('target_sharepoint_onedrive.ms_graph_client.MSGraphClient.upload_csv_file')
        self.mock_upload = self.upload_patcher.start()
        self.mock_upload.return_value = "https://test.example.com/test_file.csv"

    def teardown_method(self):
        """Clean up test method."""
        self.sharepoint_patcher.stop()
        self.upload_patcher.stop()


# TODO: Create additional tests as appropriate for your target.
