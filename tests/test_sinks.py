"""Tests for SharePoint Documents sink implementation."""

import pytest
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime
from target_sharepoint_onedrive.sinks import SharePointOneDriveSink


class TestSharePointOneDriveSink:
    """Test cases for SharePointOneDriveSink."""

    @pytest.fixture
    def mock_target(self):
        """Create a mock target instance."""
        target = Mock()
        target.config = {
            "destination_url": "https://graph.microsoft.com/v1.0/me/drive",
            "folder_path": "test-folder",
            "file_naming_scheme": "{stream_name}_{timestamp}.csv",
            "timestamp_format": "%Y%m%d_%H%M%S",
            "overwrite_files": True,
            "batch_size_rows": 1000,
        }
        return target

    @pytest.fixture
    def mock_schema(self):
        """Create a mock schema."""
        return {
            "type": "object",
            "properties": {
                "id": {"type": "integer"},
                "name": {"type": "string"},
                "value": {"type": "number"},
            }
        }

    @pytest.fixture
    def sink(self, mock_target, mock_schema):
        """Create a test sink instance."""
        with patch('target_sharepoint_onedrive.sinks.MSGraphClient') as mock_client_class:
            mock_client = Mock()
            mock_client_class.return_value = mock_client
            mock_client.destination_type = "onedrive"
            
            sink = SharePointOneDriveSink(
                target=mock_target,
                stream_name="test_stream",
                schema=mock_schema,
                key_properties=["id"]
            )
            return sink

    def test_sink_initialization(self, mock_target, mock_schema):
        """Test sink initialization."""
        with patch('target_sharepoint_onedrive.sinks.MSGraphClient') as mock_client_class:
            mock_client = Mock()
            mock_client_class.return_value = mock_client
            
            sink = SharePointOneDriveSink(
                target=mock_target,
                stream_name="test_stream",
                schema=mock_schema,
                key_properties=["id"]
            )
            
            assert sink.stream_name == "test_stream"
            assert sink.file_naming_scheme == "{stream_name}_{timestamp}.csv"
            assert sink.timestamp_format == "%Y%m%d_%H%M%S"
            assert sink.overwrite_files is True
            assert sink.records == []
            
            # Verify SharePoint client was initialized
            mock_client_class.assert_called_once_with(
                tenant_id=None,
                client_id=None,
                client_secret=None,
                destination_url="https://graph.microsoft.com/v1.0/me/drive",
                folder_path="test-folder",
            )

    def test_sink_initialization_with_credentials(self, mock_schema):
        """Test sink initialization with explicit credentials."""
        target = Mock()
        target.config = {
            "tenant_id": "test-tenant",
            "client_id": "test-client",
            "client_secret": "test-secret",
            "destination_url": "https://graph.microsoft.com/v1.0/me/drive",
            "folder_path": "test-folder",
            "file_naming_scheme": "{stream_name}_{timestamp}.csv",
            "timestamp_format": "%Y%m%d_%H%M%S",
            "overwrite_files": True,
        }
        
        with patch('target_sharepoint_onedrive.sinks.MSGraphClient') as mock_client_class:
            mock_client = Mock()
            mock_client_class.return_value = mock_client
            
            sink = SharePointOneDriveSink(
                target=target,
                stream_name="test_stream",
                schema=mock_schema,
                key_properties=["id"]
            )
            
            # Verify SharePoint client was initialized with credentials
            mock_client_class.assert_called_once_with(
                tenant_id="test-tenant",
                client_id="test-client",
                client_secret="test-secret",
                destination_url="https://graph.microsoft.com/v1.0/me/drive",
                folder_path="test-folder",
            )

    def test_start_batch(self, sink):
        """Test starting a batch."""
        # Add some records first
        sink.records = [{"id": 1, "name": "test"}]
        
        context = {"batch_id": "batch_123"}
        sink.start_batch(context)
        
        assert sink.records == []

    def test_process_record(self, sink):
        """Test processing a record."""
        record = {"id": 1, "name": "test", "value": 100}
        context = {"batch_id": "batch_123"}
        
        sink.process_record(record, context)
        
        assert len(sink.records) == 1
        assert sink.records[0] == record

    def test_process_batch_success(self, sink):
        """Test successful batch processing."""
        # Add test records
        sink.records = [
            {"id": 1, "name": "test1", "value": 100},
            {"id": 2, "name": "test2", "value": 200},
        ]

        context = {"batch_id": "batch_123"}

        # Mock the upload method
        sink.ms_graph_client.upload_csv_file.return_value = "https://onedrive.com/test_stream_20231201_120000.csv"

        sink.process_batch(context)

        # Verify upload was called with the correct data
        sink.ms_graph_client.upload_csv_file.assert_called_once()
        call_args = sink.ms_graph_client.upload_csv_file.call_args
        assert call_args[1]["data"] == [
            {"id": 1, "name": "test1", "value": 100},
            {"id": 2, "name": "test2", "value": 200},
        ]
        assert call_args[1]["overwrite"] is True
        
        # Verify records were cleared
        assert sink.records == []

    def test_process_batch_empty_records(self, sink):
        """Test batch processing with no records."""
        sink.records = []
        context = {"batch_id": "batch_123"}
        
        # Mock the upload method
        sink.ms_graph_client.upload_csv_file.return_value = "https://onedrive.com/test.csv"
        
        sink.process_batch(context)
        
        # Verify upload was not called
        sink.ms_graph_client.upload_csv_file.assert_not_called()
        
        # Verify records remain empty
        assert sink.records == []

    def test_process_batch_upload_failure(self, sink):
        """Test batch processing when upload fails."""
        # Add test records
        sink.records = [
            {"id": 1, "name": "test1", "value": 100},
        ]
        
        context = {"batch_id": "batch_123"}
        
        # Mock upload failure
        sink.ms_graph_client.upload_csv_file.side_effect = Exception("Upload failed")
        
        with pytest.raises(Exception, match="Upload failed"):
            sink.process_batch(context)
        
        # Verify records were not cleared on failure
        assert len(sink.records) == 1

    def test_generate_filename_with_timestamp(self, sink):
        """Test filename generation with timestamp."""
        context = {}
        
        # Mock datetime to get consistent timestamp
        with patch('target_sharepoint_onedrive.sinks.datetime') as mock_datetime:
            mock_datetime.now.return_value = datetime(2023, 12, 1, 12, 0, 0)
            
            filename = sink._generate_filename(context)
            
            # Should contain stream name and timestamp
            assert "test_stream" in filename
            assert "20231201_120000" in filename
            assert filename.endswith(".csv")

    def test_generate_filename_with_stream_name_only(self, sink):
        """Test filename generation with stream name only."""
        sink.file_naming_scheme = "{stream_name}.csv"
        context = {}
        
        filename = sink._generate_filename(context)
        
        assert filename == "test_stream.csv"

    def test_generate_filename_with_timestamp_only(self, sink):
        """Test filename generation with timestamp only."""
        sink.file_naming_scheme = "{timestamp}.csv"
        context = {}
        
        with patch('target_sharepoint_onedrive.sinks.datetime') as mock_datetime:
            mock_datetime.now.return_value = datetime(2023, 12, 1, 12, 0, 0)
            
            filename = sink._generate_filename(context)
            
            assert filename == "20231201_120000.csv"

    def test_generate_filename_without_csv_extension(self, sink):
        """Test filename generation without .csv extension."""
        sink.file_naming_scheme = "{stream_name}_{timestamp}"
        context = {}
        
        with patch('target_sharepoint_onedrive.sinks.datetime') as mock_datetime:
            mock_datetime.now.return_value = datetime(2023, 12, 1, 12, 0, 0)
            
            filename = sink._generate_filename(context)
            
            # Should automatically add .csv extension
            assert filename == "test_stream_20231201_120000.csv"

    def test_generate_filename_with_custom_timestamp_format(self, sink):
        """Test filename generation with custom timestamp format."""
        sink.timestamp_format = "%Y-%m-%d_%H-%M"
        sink.file_naming_scheme = "{stream_name}_{timestamp}.csv"
        context = {}
        
        with patch('target_sharepoint_onedrive.sinks.datetime') as mock_datetime:
            mock_datetime.now.return_value = datetime(2023, 12, 1, 12, 0, 0)
            
            filename = sink._generate_filename(context)
            
            assert filename == "test_stream_2023-12-01_12-00.csv"

    def test_sink_with_different_file_naming_schemes(self, mock_target, mock_schema):
        """Test sink with different file naming schemes."""
        schemes = [
            "{stream_name}.csv",
            "{timestamp}.csv",
            "{stream_name}_{timestamp}.csv",
            "data_{stream_name}_{timestamp}.csv",
        ]
        
        for scheme in schemes:
            mock_target.config["file_naming_scheme"] = scheme
            
            with patch('target_sharepoint_onedrive.sinks.MSGraphClient') as mock_client_class:
                mock_client = Mock()
                mock_client_class.return_value = mock_client
                
                sink = SharePointOneDriveSink(
                    target=mock_target,
                    stream_name="test_stream",
                    schema=mock_schema,
                    key_properties=["id"]
                )
                
                assert sink.file_naming_scheme == scheme

    def test_sink_with_different_timestamp_formats(self, mock_target, mock_schema):
        """Test sink with different timestamp formats."""
        formats = [
            "%Y%m%d_%H%M%S",
            "%Y-%m-%d_%H:%M:%S",
            "%Y%m%d",
            "%H%M%S",
        ]
        
        for fmt in formats:
            mock_target.config["timestamp_format"] = fmt
            
            with patch('target_sharepoint_onedrive.sinks.MSGraphClient') as mock_client_class:
                mock_client = Mock()
                mock_client_class.return_value = mock_client
                
                sink = SharePointOneDriveSink(
                    target=mock_target,
                    stream_name="test_stream",
                    schema=mock_schema,
                    key_properties=["id"]
                )
                
                assert sink.timestamp_format == fmt

    def test_sink_with_overwrite_disabled(self, mock_target, mock_schema):
        """Test sink with overwrite disabled."""
        mock_target.config["overwrite_files"] = False
        
        with patch('target_sharepoint_onedrive.sinks.MSGraphClient') as mock_client_class:
            mock_client = Mock()
            mock_client_class.return_value = mock_client
            
            sink = SharePointOneDriveSink(
                target=mock_target,
                stream_name="test_stream",
                schema=mock_schema,
                key_properties=["id"]
            )
            
            assert sink.overwrite_files is False
            
            # Test that overwrite=False is passed to upload
            sink.records = [{"id": 1, "name": "test", "value": 100}]
            context = {"batch_id": "batch_123"}
            sink.ms_graph_client.upload_csv_file.return_value = "https://onedrive.com/test.csv"
            
            sink.process_batch(context)
            
            # Verify overwrite=False was passed
            call_args = sink.ms_graph_client.upload_csv_file.call_args
            assert call_args[1]["overwrite"] is False
