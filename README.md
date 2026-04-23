# n8n-nodes-ms-onedrive-business

This is an n8n community node. It lets you use Microsoft OneDrive for Business and SharePoint in your n8n workflows.

Microsoft OneDrive for Business is a cloud storage service for business users that integrates with Microsoft 365 and SharePoint.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/sustainable-use-license/) workflow automation platform.

[Installation](#installation) | [Operations](#operations) | [Credentials](#credentials) | [Compatibility](#compatibility) | [Resources](#resources)

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

```bash
npm install n8n-nodes-ms-onedrive-business
```

## Operations

### Microsoft OneDrive Business Node

**File Operations:**
- **Upload** - Upload files to OneDrive/SharePoint
- **Download** - Download files as binary data
- **Get** - Retrieve file metadata
- **Delete** - Remove files
- **Rename** - Rename files
- **Search** - Search for files
- **Share** - Create sharing links (view/edit, anonymous/organization)

**Folder Operations:**
- **Create** - Create new folders
- **Delete** - Remove folders
- **Get Items** - List folder contents
- **Rename** - Rename folders
- **Search** - Search for folders
- **Share** - Create sharing links for folders

### Microsoft OneDrive Business Trigger Node

Triggers workflows when files or folders are created or updated:
- **File Created** - Trigger when new files are uploaded
- **File Updated** - Trigger when files are modified
- **Folder Created** - Trigger when new folders are created
- **Folder Updated** - Trigger when folders are modified

**Features:**
- Webhook-based real-time notifications
- Delta query for efficient change tracking
- Advanced deduplication to prevent duplicate executions
- State persistence across n8n restarts

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

## Drive Types

This node supports both:
- **User Drive**: Access individual user OneDrive (`/users/{userId}/drive`)
- **SharePoint Site Drive**: Access SharePoint site drives (`/sites/{siteId}/drive`)

Leave User ID empty to use the authenticated user's drive.

## Resources

* [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
* [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/api/resources/onedrive)
* [OneDrive API Reference](https://learn.microsoft.com/en-us/onedrive/developer/)

## Version history

### 0.1.0
- Initial release
- File operations (Upload, Download, Get, Delete, Rename, Search, Share)
- Folder operations (Create, Delete, Get Items, Rename, Search, Share)
- Trigger node with deduplication
- Support for User Drives and SharePoint Site Drives
