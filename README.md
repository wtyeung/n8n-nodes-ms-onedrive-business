# n8n-nodes-ms-onedrive-business

This is an n8n community node. It lets you use **Microsoft OneDrive for Business**, **SharePoint**, and **Excel** in your n8n workflows.

Microsoft OneDrive for Business is a cloud storage service for business users that integrates with Microsoft 365 and SharePoint.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

[Installation](#installation) | [Operations](#operations) | [File Selection](#file-selection) | [Credentials](#credentials) | [Compatibility](#compatibility) | [Resources](#resources)

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

```bash
npm install n8n-nodes-ms-onedrive-business
```

## Operations

### File Operations

- **Upload** — Upload files to OneDrive or SharePoint
- **Download** — Download files as binary data
- **Get** — Retrieve file metadata
- **Delete** — Remove files
- **Rename** — Rename files
- **Search** — Search for files by name
- **Share** — Create sharing links (view/edit, anonymous/organization)

### Folder Operations

- **Create** — Create new folders
- **Delete** — Remove folders
- **Get Items** — List folder contents
- **Rename** — Rename folders
- **Search** — Search for folders by name
- **Share** — Create sharing links for folders

### Excel Operations

Read and write data in Excel workbooks stored on OneDrive or SharePoint.

- **Read Rows** — Read rows from a worksheet, returned as structured JSON items.
  Each item includes a `_row_number` field (the actual Excel row number, e.g. `2`, `3`…) that can be used to target rows in subsequent Delete or Update operations.

- **Append or Update Row** — Write data to a worksheet.
  - **Match Record by Column** — Select a column to look up an existing row. If a match is found it is updated; if not, a new row is appended. Leave empty to always append.
  - **Matching Value** — The value to search for in the match column.
  - **Data Mode**:
    - *Auto-Map Input Data to Columns* — Automatically maps incoming `$json` fields to column headers by name.
    - *Map Each Column Manually* — Explicitly set a value (or expression) for each column.
  - Unspecified columns are **preserved** — only the columns you define are overwritten.

- **Delete Rows** — Delete one or more rows from a worksheet. Returns the deleted row data as output items.
  - **By Row Number** — Provide a single row number (use `{{ $json._row_number }}` from a Read Rows step).
  - **By Row Range** — Provide a range address such as `2:5` to delete multiple rows.

### OneDrive Business Trigger

Listens for changes in OneDrive/SharePoint using Microsoft Graph change notifications and the delta API.

#### Events

| Event | Fires when… | What to select |
|---|---|---|
| **File Created** | A new file is uploaded to the watched folder | Select the **parent folder** to watch |
| **File Updated** | A file is renamed, moved, or its content changes | Select the **parent folder** (any file inside) or a **specific file** |
| **Folder Created** | A new subfolder is created inside the watched folder | Select the **parent folder** to watch |
| **Folder Updated** | A subfolder is renamed or moved | Select the **parent folder** that contains the subfolders |

#### Important notes

- **File Updated** is the correct event for detecting file renames — renaming a file is a file metadata change, not a folder change.
- **Folder Updated** only fires when a **subfolder item itself** is renamed or moved. It does NOT fire when files inside the folder are added, renamed, or updated.
- Selecting a **specific file** for "File Updated" narrows the trigger to only that file. Selecting a **folder** watches all files inside it.
- The trigger uses Microsoft Graph subscriptions (webhook-based), so **a publicly accessible HTTPS URL is required** — `http://localhost` will not work.
- After updating the package, **deactivate and re-activate the workflow** to reinitialise the subscription and delta state.

#### Folder/File Selection for Browse mode

- **▶ FolderName** — a folder; selecting it navigates deeper
- **FileName** (no prefix) — a specific file; only available when event is "File Updated"
- **— Use Level N Folder —** — stop here and use the folder selected at Level N as the target

## File Selection

All file and folder fields support multiple selection modes:

- **Browse** (default) — Navigate up to 5 levels deep. Each level dynamically loads its contents from OneDrive/SharePoint as you select. Files and folders are displayed with icons.
- **By Path** — Type the full path directly (e.g. `/Documents/Reports/Q1.xlsx`).
- **By ID** — Provide the OneDrive item ID directly for maximum precision.
- **By Sharing Link** — Paste any OneDrive or SharePoint sharing link to access a shared file or folder directly, without needing to know its path or ID.

The same picker is used for selecting Excel workbooks, so worksheets load automatically once you select a file.

## Drive Types

Every operation supports:

- **User Drive** — The authenticated user's own OneDrive, or any other user's drive by UPN or ID. Leave User ID empty to use the authenticated user.
- **SharePoint Site Drive** — A SharePoint site's document library.
- **Shared Folder (Link)** — Paste an OneDrive or SharePoint sharing link to browse and operate on a folder that has been shared with you. Once the link is entered, the full 5-level hierarchical folder/file browser works inside the shared folder, exactly like a regular drive.

> **Note:** The "Shared Folder (Link)" drive type is supported for all **operational node** actions (download, upload, get, delete, rename, share, folder operations, Excel). It is **not supported for the trigger node** — Microsoft Graph subscriptions require direct drive ownership and do not work with sharing-link access.

## Credentials

### Prerequisites
1. Azure Active Directory (Entra ID) account
2. Permissions to create App Registrations in Azure Portal

### Setup Steps

#### 1. Create App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Enter name: `n8n OneDrive Business`
5. Select **Accounts in this organizational directory only** (or multi-tenant if needed)
6. Set Redirect URI: `https://your-n8n-instance.com/rest/oauth2-credential/callback`
7. Click **Register**

#### 2. Configure API Permissions
1. Go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `offline_access`
4. Click **Add permissions**
5. Click **Grant admin consent** (requires admin)

#### 3. Create Client Secret
1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter description: `n8n integration`
4. Select expiration (recommend 24 months)
5. Click **Add**
6. **Copy the secret value immediately** (you won't see it again)

#### 4. Get Required IDs
- **Application (client) ID**: Copy from Overview page
- **Directory (tenant) ID**: Copy from Overview page

#### 5. Configure in n8n
1. In n8n, create a new credential: **Microsoft OneDrive Business OAuth2 API**
2. Enter:
   - **Client ID**: Application (client) ID from Azure
   - **Client Secret**: Secret value from step 3
3. Click **Connect my account** and authorize

## Compatibility

- **Minimum n8n version**: 1.0.0
- **Tested with**: n8n 1.x

## Resources

* [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
* [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/api/resources/onedrive)
* [OneDrive API Reference](https://learn.microsoft.com/en-us/onedrive/developer/)

## Version history

### 0.1.12
- Fixed "Shared Folder (Link)" in trigger: now shows a clear error immediately explaining the Graph API limitation, instead of a confusing "Forbidden" message

### 0.1.11
- **By Sharing Link** file/folder selection — paste any OneDrive/SharePoint sharing link to access a shared file or folder directly
- **Shared Folder (Link)** drive type — full 5-level browse, download, upload, get, delete, rename, share, and Excel operations inside a shared folder
- Upload destination now supports "By Sharing Link" selection mode

### 0.1.3
- Added **Excel resource** with Read Rows, Append/Update Row, and Delete Rows operations
- Browser-like hierarchical folder/file picker for Excel workbook selection
- Worksheet dropdown auto-populates after selecting an Excel file
- `_row_number` field added to Read Rows output for use in Delete/Update flows
- Delete Rows supports Row Number mode (via `_row_number`) and Row Range mode
- Append/Update preserves existing column values — only specified columns are overwritten
- Match Record by Column is a selectable dropdown with hint text

### 0.1.2
- Extended Microsoft OAuth2 credential with credential test
- Simplified credential to essential fields only

### 0.1.1
- Hierarchical folder browse UI for all folder/file selection fields

### 0.1.0
- Initial release
- File operations (Upload, Download, Get, Delete, Rename, Search, Share)
- Folder operations (Create, Delete, Get Items, Rename, Search, Share)
- Support for User Drives and SharePoint Site Drives
