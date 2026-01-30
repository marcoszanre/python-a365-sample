# Agent 365 MCP Servers - Complete Reference

This document provides a comprehensive guide to all available MCP (Model Context Protocol) servers in the Agent 365 ecosystem, based on the official Microsoft documentation.

## Overview

Agent 365 tooling servers are enterprise-grade MCP servers that give agents safe, governed access to business systems such as Microsoft Outlook, Microsoft Teams, Microsoft SharePoint and OneDrive, Microsoft Dataverse, and more through the tooling gateway.

## Available MCP Servers

| MCP Server | Server ID | Scope | Description |
|------------|-----------|-------|-------------|
| **Outlook Mail** | `mcp_MailTools` | `McpServers.Mail.All` | Email operations (create, send, reply, search, delete) |
| **Outlook Calendar** | `mcp_CalendarTools` | `McpServers.Calendar.All` | Calendar events (create, update, accept/decline, find meeting times) |
| **Teams** | `mcp_TeamsServer` | `McpServers.Teams.All` | Chats, channels, messages, and team management |
| **Word** | `mcp_WordServer` | `McpServers.Word.All` | Create documents, read content, manage comments |
| **SharePoint/OneDrive** | `mcp_ODSPRemoteServer` | `McpServers.OneDriveSharepoint.All` | Files, folders, sharing, sites, document libraries |
| **SharePoint Lists** | `mcp_SharePointListsTools` | `McpServers.SharepointLists.All` | Lists, columns, and list items CRUD |
| **User Profile** | `mcp_MeServer` | `McpServers.Me.All` | User profiles, managers, direct reports, org search |
| **Copilot Search** | `mcp_M365Copilot` | `McpServers.CopilotMCP.All` | Search across all M365 content via Copilot |
| **Knowledge** | `mcp_KnowledgeTools` | `McpServers.Knowledge.All` | Knowledge management tools |
| **Admin365 Graph** | `mcp_Admin365_GraphTools` | `McpServers.Admin365Graph.All` | Admin Graph operations |
| **M365 Admin** | `mcp_AdminTools` | `McpServers.M365Admin.All` | M365 admin operations |
| **DA Search** | `mcp_DASearch` | `McpServers.DASearch.All` | Data and analytics search |

---

## Detailed Server Reference

### 📧 mcp_MailTools - Outlook Mail

**Scope:** `McpServers.Mail.All`

**Available Tools:**
- `mcp_MailTools_graph_mail_createMessage` - Create a draft email
- `mcp_MailTools_graph_mail_sendMail` - Send an email
- `mcp_MailTools_graph_mail_sendDraft` - Send an existing draft
- `mcp_MailTools_graph_mail_getMessage` - Get a message by ID
- `mcp_MailTools_graph_mail_searchMessages` - Search emails with KQL queries
- `mcp_MailTools_graph_mail_reply` - Reply to a message
- `mcp_MailTools_graph_mail_replyAll` - Reply-all to a message
- `mcp_MailTools_graph_mail_updateMessage` - Update message properties
- `mcp_MailTools_graph_mail_deleteMessage` - Delete a message
- `mcp_MailTools_graph_mail_listSent` - List sent messages

**Use Cases:**
- Create and send emails (HTML/plain text)
- Reply and reply-all to messages
- Search emails with KQL queries
- Manage drafts and sent items

**Key Features:**
- HTML and plain text support
- To, CC, and BCC recipients
- KQL-style message search
- OData query support
- Draft management

---

### 📅 mcp_CalendarTools - Outlook Calendar

**Scope:** `McpServers.Calendar.All`

**Available Tools:**
- `mcp_CalendarTools_graph_createEvent` - Create a calendar event
- `mcp_CalendarTools_graph_getEvent` - Get an event by ID
- `mcp_CalendarTools_graph_updateEvent` - Update an event
- `mcp_CalendarTools_graph_deleteEvent` - Delete an event
- `mcp_CalendarTools_graph_listEvents` - List calendar events
- `mcp_CalendarTools_graph_listCalendarView` - Get events in a time range
- `mcp_CalendarTools_graph_acceptEvent` - Accept an invitation
- `mcp_CalendarTools_graph_declineEvent` - Decline an invitation
- `mcp_CalendarTools_graph_cancelEvent` - Cancel an event
- `mcp_CalendarTools_graph_findMeetingTimes` - Find available meeting times
- `mcp_CalendarTools_graph_getSchedule` - Get free/busy schedule

**Use Cases:**
- Create/update calendar events
- Accept/decline invitations
- Find available meeting times
- Get free/busy schedules
- Create Teams/Skype online meetings

**Key Features:**
- Recurring event support
- Online meeting integration (Teams, Skype)
- Attendee management (required, optional, resource)
- Time zone support
- Availability checking

---

### 💬 mcp_TeamsServer - Microsoft Teams

**Scope:** `McpServers.Teams.All`

**Available Tools:**

**Chat Operations:**
- `mcp_graph_chat_createChat` - Create a new chat (1:1 or group)
- `mcp_graph_chat_getChat` - Get chat details
- `mcp_graph_chat_listChats` - List user's chats
- `mcp_graph_chat_updateChat` - Update chat properties (topic)
- `mcp_graph_chat_deleteChat` - Delete a chat
- `mcp_graph_chat_postMessage` - Post a message to chat
- `mcp_graph_chat_getChatMessage` - Get a specific message
- `mcp_graph_chat_listChatMessages` - List messages in a chat
- `mcp_graph_chat_updateChatMessage` - Edit a message
- `mcp_graph_chat_deleteChatMessage` - Soft-delete a message
- `mcp_graph_chat_addChatMember` - Add member to chat
- `mcp_graph_chat_listChatMembers` - List chat members

**Channel & Team Operations:**
- `mcp_graph_teams_createChannel` - Create a channel
- `mcp_graph_teams_createPrivateChannel` - Create a private channel
- `mcp_graph_teams_getChannel` - Get channel details
- `mcp_graph_teams_listChannels` - List channels in a team
- `mcp_graph_teams_updateChannel` - Update channel properties
- `mcp_graph_teams_postChannelMessage` - Post to a channel
- `mcp_graph_teams_replyToChannelMessage` - Reply in a thread
- `mcp_graph_teams_listChannelMessages` - List channel messages
- `mcp_graph_teams_addChannelMember` - Add member to channel
- `mcp_graph_teams_listChannelMembers` - List channel members
- `mcp_graph_teams_updateChannelMember` - Update member role
- `mcp_graph_teams_getTeam` - Get team details
- `mcp_graph_teams_listTeams` - List user's joined teams

**Use Cases:**
- Create 1:1 and group chats
- Send/edit/delete messages
- Manage channels (standard, private, shared)
- Add members to chats/channels
- Post to channels and reply to threads

**Key Features:**
- Full CRUD for chats and channels
- Support for private, shared, and standard channels
- Message threading and replies
- Member role management (owner, member)
- OData query support for filtering

---

### 📄 mcp_WordServer - Microsoft Word

**Scope:** `McpServers.Word.All`

**Available Tools:**
- `WordCreateNewDocument` - Create a new Word document in OneDrive
- `WordGetDocumentContent` - Fetch document content from SharePoint/OneDrive URL
- `WordCreateNewComment` - Add a comment to a document
- `WordReplyToComment` - Reply to an existing comment

**Use Cases:**
- Create Word documents in OneDrive
- Read document content from SharePoint/OneDrive URLs
- Add and reply to comments
- Extract plain text from DOCX files

**Key Features:**
- HTML and plain text content support
- Auto-generated file names with timestamp
- Comment thread management
- Returns Microsoft Graph DriveItem metadata

---

### 📁 mcp_ODSPRemoteServer - SharePoint & OneDrive

**Scope:** `McpServers.OneDriveSharepoint.All`

**Available Tools:**
- `createFolder` - Create a folder in document library
- `findFileOrFolder` - Search for files/folders
- `findSite` - Find SharePoint sites
- `createSmallTextFile` - Create/upload text files (<5MB)
- `readSmallTextFile` - Read/download text files
- `moveSmallFile` - Move files (<5MB, same site)
- `deleteFileOrFolder` - Delete files/folders
- `renameFileOrFolder` - Rename files/folders
- `shareFileOrFolder` - Share with role assignments
- `getFileOrFolderMetadata` - Get metadata by ID
- `getFileOrFolderMetadataByUrl` - Get metadata by URL
- `getFolderChildren` - List folder contents (top 20)
- `listDocumentLibrariesInSite` - List document libraries
- `getDefaultDocumentLibraryInSite` - Get default library
- `setSensitivityLabelOnFile` - Apply sensitivity labels

**Use Cases:**
- Create/manage files and folders
- Search for files and sites
- Share files with roles (read/write)
- Apply sensitivity labels
- List document libraries

**Key Features:**
- OneDrive and SharePoint Online support
- Automatic naming conflict resolution
- Sharing with custom messages
- Sensitivity label compliance
- DriveItem abstraction for files/folders

---

### 📋 mcp_SharePointListsTools - SharePoint Lists

**Scope:** `McpServers.SharepointLists.All`

**Available Tools:**
- `sharepoint_createList` - Create a new SharePoint list
- `sharepoint_listLists` - List all lists on a site
- `sharepoint_createListColumn` - Create a column
- `sharepoint_editListColumn` - Edit a column
- `sharepoint_deleteListColumn` - Delete a column
- `sharepoint_listListColumns` - List columns in a list
- `sharepoint_createListItem` - Create a list item
- `sharepoint_updateListItem` - Update a list item
- `sharepoint_deleteListItem` - Delete a list item
- `sharepoint_listListItems` - List items with filtering
- `sharepoint_searchSitesByName` - Search sites by name
- `sharepoint_getSiteByPath` - Resolve site by path
- `sharepoint_listSubsites` - List child sites
- `sharepoint_search` - KQL search for sites

**Supported Column Types:**
- Text (single-line, multiline, plain, rich)
- Number (with decimal, min/max)
- Choice (checkbox, dropdown, radio)
- Boolean
- DateTime (date only, date and time)
- Person or Group (single, multiple)
- Lookup (reference other lists)
- Hyperlink or Picture

**Use Cases:**
- Create and manage SharePoint lists
- Add/edit/delete columns
- CRUD operations on list items
- Search and discover sites

---

### 👤 mcp_MeServer - User Profile

**Scope:** `McpServers.Me.All`

**Available Tools:**
- `mcp_graph_getMyProfile` - Get current user's profile
- `mcp_graph_getMyManager` - Get current user's manager
- `mcp_graph_getUserProfile` - Get any user's profile
- `mcp_graph_getUsersManager` - Get any user's manager
- `mcp_graph_getDirectReports` - Get user's direct reports
- `mcp_graph_listUsers` - Search/list users in org

**Use Cases:**
- Get current user's profile and manager
- Look up any user's profile
- Navigate org hierarchy (manager, direct reports)
- Search users in the organization

**Key Features:**
- Self-knowledge (signed-in user context)
- Organizational hierarchy navigation
- Free-text search with automatic fallback
- OData filtering and pagination

**Important Notes:**
- Use `getMyProfile` for signed-in user, NOT `getUserProfile` with 'me'
- userIdentifier must be object ID (GUID) or userPrincipalName (UPN)
- If only display name available, use `listUsers` to look up first

---

### 🔍 mcp_M365Copilot - Copilot Search

**Scope:** `McpServers.CopilotMCP.All`

**Available Tools:**
- `Copilot Chat` - Search across Microsoft 365 ecosystem

**Use Cases:**
- Search across all Microsoft 365 content
- Find documents, emails, Teams chats, SharePoint sites
- Ground responses with specific files
- Fallback when workload-specific tools unavailable
- Multi-turn conversations with context retention

**Key Features:**
- Direct Microsoft 365 Copilot integration
- Persistent conversation IDs
- File grounding for contextual responses
- Location and time zone context support

**When to Use:**
> Use this tool for any user request that might require finding, searching, discovering, or locating information contained within Microsoft 365 content—including documents, PDFs, spreadsheets, emails, sites, reports, or files.

---

## CLI Commands

### Discover Available Servers
```powershell
a365 develop list-available
```

### Add MCP Servers
```powershell
a365 develop add-mcp-servers mcp_MailTools mcp_CalendarTools mcp_TeamsServer
```

### List Configured Servers
```powershell
a365 develop list-configured
```

### Remove MCP Servers
```powershell
a365 develop remove-mcp-servers mcp_MailTools
```

---

## Documentation Links

For detailed tool parameters and examples, refer to the official Microsoft documentation:

- [Tooling Overview](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview)
- [Mail Tools Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/mail)
- [Calendar Tools Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/calendar)
- [Teams Tools Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/teams)
- [Word Tools Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/word)
- [SharePoint/OneDrive Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/odspremoteserver)
- [SharePoint Lists Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/sharepointlisttools)
- [User Profile (Me) Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/me)
- [Copilot Search Reference](https://learn.microsoft.com/en-us/microsoft-agent-365/mcp-server-reference/searchtools)

---

## Notes

- All MCP servers require appropriate permissions granted by tenant admin
- Authentication is automatically configured via `ToolingManifest.json`
- Each server is governed through Microsoft 365 admin center
- Full observability via Microsoft Defender for auditing
- All operations are traceable and compliant with organizational policies
