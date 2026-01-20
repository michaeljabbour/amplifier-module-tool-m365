# amplifier-module-tool-m365

Microsoft 365 collaboration module for Amplifier. Provides integration with Teams, SharePoint, Outlook, and Planner via Microsoft Graph API.

## Features

- **Microsoft Teams** - Post/read channel messages via webhooks and API
- **SharePoint** - Upload/download/list documents
- **Outlook** - Send email notifications
- **Planner** - List tasks

## Installation

```bash
pip install amplifier-module-tool-m365
```

Or from source:
```bash
pip install git+https://github.com/michaeljabbour/amplifier-module-tool-m365
```

## Configuration

### Required Environment Variables

```bash
export M365_TENANT_ID="your-tenant-id"
export M365_CLIENT_ID="your-client-id"
export M365_CLIENT_SECRET="your-client-secret"
```

### Optional: Teams Webhooks (Recommended for Posting)

```bash
export M365_TEAMS_WEBHOOKS="general=https://outlook.office.com/webhook/...,alerts=https://...,handoffs=https://..."
```

## Azure AD App Setup

1. Go to https://portal.azure.com → Azure Active Directory → App registrations
2. New registration → Name your app
3. API permissions → Add:
   - `User.Read.All` (Application)
   - `Group.Read.All` (Application)
   - `ChannelMessage.Read.All` (Application)
   - `Sites.ReadWrite.All` (Application)
   - `Mail.Send` (Application)
   - `Tasks.ReadWrite.All` (Application)
4. Grant admin consent
5. Certificates & secrets → New client secret
6. Copy: Application (client) ID, Directory (tenant) ID, Client secret value

## Usage

### With Amplifier Bundle

```yaml
# In your bundle
tools:
  - module: tool-collab-core
    source: git+https://github.com/michaeljabbour/amplifier-module-tool-collab-core
  - module: tool-m365
    source: git+https://github.com/michaeljabbour/amplifier-module-tool-m365
```

### Direct Python Usage

```python
import asyncio
from amplifier_module_tool_m365 import M365Provider

async def main():
    provider = M365Provider()
    
    # List users
    users = await provider.list_users(limit=5)
    for user in users:
        print(f"  {user.display_name} ({user.email})")
    
    # Post to Teams channel
    await provider.post_message("general", "Hello from Amplifier!", "Status Update")
    
    # Upload to SharePoint
    doc = await provider.upload_document(
        "report.md",
        "# Report\n\nContent here...",
        folder_path="Amplifier/reports"
    )
    print(f"Uploaded: {doc.web_url}")

asyncio.run(main())
```

### Via Core Module

```python
from amplifier_module_tool_collab_core import get_provider
import amplifier_module_tool_m365  # Registers 'm365' provider

provider = get_provider("m365")
await provider.post_message("general", "Hello!")
```

## Capabilities

| Feature | Method | Notes |
|---------|--------|-------|
| List users | `list_users()` | Azure AD users |
| Get user | `get_user(id)` | By ID or UPN |
| List channels | `list_channels()` | Teams channels |
| Read messages | `get_messages()` | Requires team_id |
| Post message | `post_message()` | Via webhook (preferred) |
| List documents | `list_documents()` | SharePoint |
| Upload document | `upload_document()` | SharePoint |
| Download document | `download_document()` | SharePoint |
| List tasks | `list_tasks()` | Planner |
| Send email | `send_email()` | Outlook |

## License

MIT
