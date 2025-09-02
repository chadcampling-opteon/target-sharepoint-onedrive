# target-sharepoint-onedrive

`target-sharepoint-onedrive` is a Singer target for SharePoint Document Libraries and OneDrive that writes records as CSV files using the Microsoft Graph API.

Built with the [Meltano Target SDK](https://sdk.meltano.com).

## Features

- **SharePoint Integration**: Uploads CSV files directly to SharePoint document libraries
- **OneDrive Integration**: Uploads CSV files to OneDrive for Business or personal OneDrive
- **Microsoft Graph API**: Uses the official Microsoft Graph API for reliable file operations
- **Batch Processing**: Processes records in configurable batches for optimal performance
- **Flexible File Naming**: Customizable file naming schemes with timestamp and stream name variables
- **Stream Maps**: Transform stream names and data using stream maps
- **Schema Flattening**: Optionally flatten nested JSON structures
- **Large File Support**: Handles files larger than 4MB using upload sessions
- **Simple Configuration**: Uses SharePoint or OneDrive destination URLs for easy setup
- **Flexible Authentication**: Supports both explicit credentials and Azure default credential chain

## Installation

Install from PyPI:

```bash
pipx install target-sharepoint-onedrive
```

Install from GitHub:

```bash
pipx install git+https://github.com/chadcampling-opteon/target-sharepoint-onedrive.git@main
```

## Configuration

### Prerequisites

Before using this target, you need to:

1. **Register an Entra ID Application** (for explicit authentication):
   - Go to Azure Portal > Azure Active Directory > App registrations
   - Create a new registration
   - Note down the Application (client) ID and Directory (tenant) ID
   - Create a client secret and note it down

2. **Configure API Permissions**:
   - In your app registration, go to API permissions
   - Add Microsoft Graph permissions:
     - `Sites.ReadWrite.All` (for SharePoint site access)
     - `Files.ReadWrite.All` (for file operations)
     - `User.Read.All` (for accessing other users' OneDrive)

3. **Get SharePoint or OneDrive Destination URL**:
   - For SharePoint: Get your SharePoint site URL from your browser and convert it to Graph API format
   - For OneDrive: Use the appropriate OneDrive URL format (see Destination URL section below)

### Authentication Options

This target supports two authentication methods:

#### 1. Explicit Credentials
Provide `tenant_id`, `client_id`, and `client_secret` in your configuration.

#### 2. Azure Default Credential Chain
For access, you can omit the authentication fields and rely on Azure's default credential chain, which will automatically use:
- Environment variables (`AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID`)
- Managed Identity (when running in Azure)
- Azure CLI credentials
- Visual Studio Code credentials
- Azure PowerShell credentials

**Example OneDrive configuration without explicit credentials:**
```json
{
  "destination_url": "https://graph.microsoft.com/v1.0/me/drive",
  "folder_path": "data/exports"
}
```

### Configuration Options

| Setting | Required | Default | Description |
|---------|----------|---------|-------------|
| `tenant_id` | No | - | Azure AD tenant ID, required if not using DefaultAzureCredential resolution |
| `client_id` | No | - | Azure AD application client ID, required if not using DefaultAzureCredential resolution |
| `client_secret` | No | - | Azure AD application client secret, required if not using DefaultAzureCredential resolution |
| `destination_url` | Yes | - | SharePoint/OneDrive destination URL |
| `folder_path` | No | "" | Optional folder path within the document library |
| `file_naming_scheme` | No | `{stream_name}_{timestamp}.csv` | Scheme for naming output files, replacement options for stream_name, timestamp, batch_id|
| `timestamp_format` | No | `%Y%m%d_%H%M%S` | Format for timestamp in filename |
| `overwrite_files` | No | `true` | Whether to overwrite existing files |
| `batch_size_rows` | No | `10000` | Maximum number of records per file |
| `flattening_enabled` | No | `false` | Enable schema flattening |
| `flattening_max_depth` | No | `0` | Max depth to flatten schemas |
| `stream_maps` | No | - | Config object for stream maps |
| `stream_map_config` | No | - | User-defined config for stream maps |

### Configuration Setup

Use a SharePoint or OneDrive destination URL to automatically resolve site and drive information:

#### SharePoint Configuration Example:
```json
{
  "tenant_id": "your-tenant-id",
  "client_id": "your-client-id",
  "client_secret": "your-client-secret",
  "destination_url": "https://graph.microsoft.com/v1.0/sites/yourdomain.sharepoint.com:/sites/YourSiteName:/",
  "folder_path": "data/exports"
}
```

#### OneDrive Configuration Examples:

**Personal OneDrive (current user):**
```json
{
  "destination_url": "https://graph.microsoft.com/v1.0/me/drive",
  "folder_path": "data/exports"
}
```

**OneDrive for Business (specific user):**
```json
{
  "destination_url": "https://graph.microsoft.com/v1.0/users/user@yourdomain.com/drive",
  "folder_path": "data/exports"
}
```

**OneDrive by User ID:**
```json
{
  "destination_url": "https://graph.microsoft.com/v1.0/users/12345678-1234-1234-1234-123456789012/drive",
  "folder_path": "data/exports"
}
```

**Direct Drive Access (if you know the drive ID):**
```json
{
  "destination_url": "https://graph.microsoft.com/v1.0/drives/b!1234567890123456789012345678901234567890123456789012345678901234",
  "folder_path": "data/exports"
}
```

### Destination URL Formats

#### SharePoint URLs:
```
https://graph.microsoft.com/v1.0/sites/{domain}.sharepoint.com:/sites/{site-name}:/
```

#### OneDrive URLs:
- **Personal OneDrive**: `https://graph.microsoft.com/v1.0/me/drive`
- **User OneDrive**: `https://graph.microsoft.com/v1.0/users/{user-email-or-id}/drive`
- **Direct Drive**: `https://graph.microsoft.com/v1.0/drives/{drive-id}`

The target will automatically:
1. Resolve the site/drive ID from the URL
2. Find the appropriate document library or OneDrive root
3. Use that for file uploads

### File Naming Variables

The `file_naming_scheme` supports the following variables:

- `{stream_name}`: Name of the stream
- `{timestamp}`: Current timestamp in the specified format
- `{batch_id}`: Unique batch identifier

Example: `{stream_name}_{timestamp}_batch_{batch_id}.csv`

### Sample Configuration

See [config_sample.json](config_sample.json) for a complete SharePoint configuration example and [config_sample_onedrive.json](config_sample_onedrive.json) for OneDrive examples.

### Configure using environment variables

This Singer target will automatically import any environment variables within the working directory's
`.env` if the `--config=ENV` is provided, such that config values will be considered if a matching
environment variable is set either in the terminal context or in the `.env` file.

Example `.env` file for SharePoint:
```env
TARGET_SHAREPOINT_ONEDRIVE_TENANT_ID=your-tenant-id
TARGET_SHAREPOINT_ONEDRIVE_CLIENT_ID=your-client-id
TARGET_SHAREPOINT_ONEDRIVE_CLIENT_SECRET=your-client-secret
TARGET_SHAREPOINT_ONEDRIVE_DESTINATION_URL=https://graph.microsoft.com/v1.0/sites/yourdomain.sharepoint.com:/sites/YourSiteName:/
```

Example `.env` file for OneDrive:
```env
TARGET_SHAREPOINT_ONEDRIVE_DESTINATION_URL=https://graph.microsoft.com/v1.0/me/drive
TARGET_SHAREPOINT_ONEDRIVE_FOLDER_PATH=data/exports
```

**Note**: For OneDrive with environment-based authentication (managed identity, service principal, etc.), you may not need to set the authentication environment variables explicitly.

## Usage

You can easily run `target-sharepoint-onedrive` by itself or in a pipeline using [Meltano](https://meltano.com/).

### Executing the Target Directly

```bash
target-sharepoint-onedrive --version
target-sharepoint-onedrive --help
# Test using the "Smoke Test" tap:
tap-smoke-test | target-sharepoint-onedrive --config config.json
```

### Using with Meltano

1. **Install Meltano** (if you haven't already):
```bash
pipx install meltano
```

2. **Initialize Meltano** within your project:
```bash
cd your-project
meltano install
```

3. **Configure the target**:

**For SharePoint:**
```bash
meltano config target-sharepoint-onedrive set tenant_id your-tenant-id
meltano config target-sharepoint-onedrive set client_id your-client-id
meltano config target-sharepoint-onedrive set client_secret your-client-secret
meltano config target-sharepoint-onedrive set destination_url "https://graph.microsoft.com/v1.0/sites/yourdomain.sharepoint.com:/sites/YourSiteName:/"
```

**For OneDrive:**
```bash
meltano config target-sharepoint-onedrive set destination_url "https://graph.microsoft.com/v1.0/me/drive"
meltano config target-sharepoint-onedrive set folder_path "data/exports"
```

4. **Run an ELT pipeline**:
```bash
meltano run tap-smoke-test target-sharepoint-onedrive
```

## Getting Destination URLs

### SharePoint Destination URL

1. **Get your SharePoint site URL** from your browser when accessing the site
2. **Convert to Graph API format**:
   - Original: `https://yourdomain.sharepoint.com/sites/YourSiteName`
   - Graph API: `https://graph.microsoft.com/v1.0/sites/yourdomain.sharepoint.com:/sites/YourSiteName:/`

**Example:**
- Original URL: `https://opteonpropertygroup.sharepoint.com/sites/CAMATeamWorkItems`
- Destination URL: `https://graph.microsoft.com/v1.0/sites/opteonpropertygroup.sharepoint.com:/sites/CAMATeamWorkItems:/`

### OneDrive Destination URLs

#### Personal OneDrive (Current User)
Use this URL to access the current user's OneDrive:
```
https://graph.microsoft.com/v1.0/me/drive
```

#### OneDrive for Business (Specific User)
To access a specific user's OneDrive, you can use their email address or user ID:

**By Email:**
```
https://graph.microsoft.com/v1.0/users/user@yourdomain.com/drive
```

**By User ID:**
```
https://graph.microsoft.com/v1.0/users/12345678-1234-1234-1234-123456789012/drive
```

#### Finding User IDs
To find a user's ID, you can:
1. Use the Microsoft Graph Explorer: `https://developer.microsoft.com/en-us/graph/graph-explorer`
2. Query: `GET https://graph.microsoft.com/v1.0/users`
3. Look for the `id` field in the response

#### Direct Drive Access
If you know the drive ID, you can access it directly:
```
https://graph.microsoft.com/v1.0/drives/{drive-id}
```

**Note**: For OneDrive access, ensure your Azure AD application has the appropriate permissions:
- `Files.ReadWrite.All` (for file operations)
- `User.Read.All` (if accessing other users' OneDrive)

## Developer Resources

Follow these instructions to contribute to this project.

### Initialize your Development Environment

```bash
pipx install poetry
poetry install
```

### Create and Run Tests

Create tests within the `target_sharepoint_onedrive/tests` subfolder and then run:

```bash
poetry run pytest
```

You can also test the `target-sharepoint-onedrive` CLI interface directly using `poetry run`:

```bash
poetry run target-sharepoint-onedrive --help
```

### Testing with Meltano

_**Note:** This target will work in any Singer environment and does not require Meltano. Examples here are for convenience and to streamline end-to-end orchestration scenarios._

Next, install Meltano (if you haven't already) and any needed plugins:

```bash
# Install meltano
pipx install meltano
# Initialize meltano within this directory
cd target-sharepoint-onedrive
meltano install
```

Now you can test and orchestrate using Meltano:

```bash
# Test invocation:
meltano invoke target-sharepoint-onedrive --version
# OR run a test `elt` pipeline:
meltano elt tap-smoke-test target-sharepoint-onedrive
```

### SDK Dev Guide

See the dev guide for more instructions on how to use the SDK to develop your own taps and targets.

## About

Singer.io compatible Target for SharePoint Document Libraries. Created with the Meltano SDK.

### Resources

- [Meltano Target SDK](https://sdk.meltano.com)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [SharePoint REST API](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
