"""SharePoint OneDrive target sink class, which handles writing streams."""

from __future__ import annotations

import logging
from datetime import datetime
import os
from typing import Any, Dict

from singer_sdk.sinks import BatchSink

from target_sharepoint_onedrive.ms_graph_client import MSGraphClient

logger = logging.getLogger(__name__)


class SharePointOneDriveSink(BatchSink):
    """SharePoint OneDrive target sink class."""

    def __init__(self, target: Any, stream_name: str, schema: Dict, key_properties: list[str] | None) -> None:
        """Initialize the sink.

        Args:
            target: The target instance
            stream_name: Name of the stream
            schema: JSON schema for the stream
            key_properties: Primary key properties for the stream
        """
        super().__init__(target, stream_name, schema, key_properties)
        
        # Get configuration
        config = self.config
        
        # Initialize SharePoint client with destination URL (type auto-detected)
        self.ms_graph_client = MSGraphClient(
            tenant_id=config.get("tenant_id"),
            client_id=config.get("client_id"),
            client_secret=config.get("client_secret"),
            destination_url=config["destination_url"],
            folder_path=config.get("folder_path", ""),
        )
        
        # Get file configuration
        self.file_naming_scheme = config.get("file_naming_scheme", "{stream_name}_{timestamp}.csv")
        self.timestamp_format = config.get("timestamp_format", "%Y%m%d_%H%M%S")
        self.overwrite_files = config.get("overwrite_files", True)
        
        # Initialize batch storage
        self.records: list[Dict[str, Any]] = []

    def start_batch(self, context: dict) -> None:
        """Start a batch.

        Args:
            context: Stream partition or context dictionary.
        """
        self.records = []
        logger.info(f"Starting batch for stream: {self.stream_name}")

    def process_record(self, record: dict, context: dict) -> None:
        """Process the record.

        Args:
            record: Individual record in the stream.
            context: Stream partition or context dictionary.
        """
        self.records.append(record)

    def process_batch(self, context: dict) -> None:
        """Write out any prepped records and return once fully written.

        Args:
            context: Stream partition or context dictionary.
        """
        if not self.records:
            logger.info(f"No records to process for stream: {self.stream_name}")
            return
        
        try:
            # Generate filename
            filename = self._generate_filename(context)
            
            # Upload to SharePoint or OneDrive
            file_url = self.ms_graph_client.upload_csv_file(
                data=self.records,
                filename=filename,
                overwrite=self.overwrite_files,
            )
            
            destination_name = "OneDrive" if self.ms_graph_client.destination_type == "onedrive" else "SharePoint"
            logger.info(
                f"Successfully uploaded {len(self.records)} records to {destination_name}: {file_url}"
            )
            
            # Clear records for next batch
            self.records = []
            
        except Exception as e:
            logger.error(f"Failed to upload batch for stream {self.stream_name}: {str(e)}")
            raise

    def _generate_filename(self, context: dict) -> str:
        """Generate filename based on naming scheme.

        Args:
            context: Stream partition or context dictionary.

        Returns:
            Generated filename
        """
        timestamp = datetime.now().strftime(self.timestamp_format)
        batch_id = context.get("batch_id", "unknown")
        
        filename = self.file_naming_scheme.format(
            stream_name=self.stream_name,
            timestamp=timestamp,
            batch_id=batch_id,
        )
        
        # Ensure filename ends with .csv
        if not filename.endswith(".csv"):
            filename += ".csv"
        
        return filename
