# Undocumented MCP Servers

> **Last Updated:** January 30, 2026
>
> **Note:** This document covers MCP servers that are available in the Agent 365 catalog but do not yet have official Microsoft documentation. These servers are likely in early preview stages within the Frontier program.

## Overview

The following MCP servers are returned by `a365 develop list-available` but do not have dedicated reference documentation on Microsoft Learn. Use these servers with caution as their APIs may change without notice.

---

## Undocumented Servers

### 1. mcp_DASearch - Data & Analytics Search

| Property | Value |
|----------|-------|
| **Server ID** | `mcp_DASearch` |
| **URL** | `https://agent365.svc.cloud.microsoft/agents/servers/mcp_DASearch` |
| **Required Scope** | `McpServers.DASearch.All` |
| **Audience** | `ea9ffc3e-8a23-4a7d-836d-234d7c7565c1` |
| **Documentation** | ❌ Not available |

#### Speculated Capabilities
Based on the naming convention, this server likely provides:
- Data analytics search capabilities
- Integration with Microsoft data and analytics services
- Possibly Power BI or data warehouse search functionality
- Enterprise data discovery tools

#### Usage (Untested)
```powershell
# Add to your agent
a365 develop add-mcp-servers mcp_DASearch
```

---

### 2. mcp_TeamsCanaryServer - Teams Canary/Preview

| Property | Value |
|----------|-------|
| **Server ID** | `mcp_TeamsCanaryServer` |
| **URL** | `https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsCanaryServer` |
| **Required Scope** | `McpServers.Teams.All` |
| **Audience** | `ea9ffc3e-8a23-4a7d-836d-234d7c7565c1` |
| **Documentation** | ❌ Not available |

#### Speculated Capabilities
This server uses the **same scope** as `mcp_TeamsServer`, suggesting it's a **canary/preview build** of the Teams MCP server:
- Early access to new Teams MCP features
- Testing ground for upcoming Teams integrations
- Same tools as `mcp_TeamsServer` but with experimental features
- May have breaking changes more frequently

#### ⚠️ Recommendation
For production use, prefer `mcp_TeamsServer` which is fully documented. Use `mcp_TeamsCanaryServer` only for:
- Testing upcoming features
- Providing feedback to Microsoft
- Development/staging environments

#### Usage (Untested)
```powershell
# Add to your agent (use with caution)
a365 develop add-mcp-servers mcp_TeamsCanaryServer
```

---

### 3. mcp_KnowledgeTools - Knowledge Management

| Property | Value |
|----------|-------|
| **Server ID** | `mcp_KnowledgeTools` |
| **URL** | `https://agent365.svc.cloud.microsoft/agents/servers/mcp_KnowledgeTools` |
| **Required Scope** | `McpServers.Knowledge.All` |
| **Audience** | `ea9ffc3e-8a23-4a7d-836d-234d7c7565c1` |
| **Documentation** | ❌ Not available |

#### Speculated Capabilities
Based on the naming and Microsoft's knowledge management ecosystem, this server likely provides:
- Integration with Microsoft Viva Topics
- Knowledge base search and retrieval
- Topic extraction and management
- Organizational knowledge graph access
- Expert finder functionality
- Learning content discovery (Viva Learning integration)

#### Potential Use Cases
- Building agents that can tap into organizational knowledge
- Creating intelligent assistants that understand company topics
- Automating knowledge curation workflows
- Connecting agents to learning and training content

#### Usage (Untested)
```powershell
# Add to your agent
a365 develop add-mcp-servers mcp_KnowledgeTools
```

---

### 4. mcp_Admin365_GraphTools - Admin Graph Operations

| Property | Value |
|----------|-------|
| **Server ID** | `mcp_Admin365_GraphTools` |
| **URL** | `https://agent365.svc.cloud.microsoft/agents/servers/mcp_Admin365_GraphTools` |
| **Required Scope** | `McpServers.Admin365Graph.All` |
| **Audience** | `ea9ffc3e-8a23-4a7d-836d-234d7c7565c1` |
| **Documentation** | ❌ Not available |

#### Speculated Capabilities
This server likely provides administrative Microsoft Graph operations:
- User and group management
- License assignment and management
- Directory operations
- Tenant configuration
- Security and compliance settings
- Audit log access
- Service health monitoring

#### ⚠️ Important Security Note
This server likely requires **elevated admin privileges**. Tools may include:
- Creating/deleting users
- Modifying group memberships
- Changing tenant settings
- Accessing sensitive audit data

**Use with extreme caution** and ensure proper governance is in place.

#### Potential Use Cases
- IT admin automation agents
- Self-service user management bots
- Compliance and audit agents
- License optimization assistants

#### Usage (Untested)
```powershell
# Add to your agent (requires admin privileges)
a365 develop add-mcp-servers mcp_Admin365_GraphTools
```

---

## Comparison: Documented vs Undocumented

| Server | Scope | Documented | Production Ready |
|--------|-------|------------|------------------|
| `mcp_MailTools` | `McpServers.Mail.All` | ✅ Yes | ✅ Yes |
| `mcp_CalendarTools` | `McpServers.Calendar.All` | ✅ Yes | ✅ Yes |
| `mcp_TeamsServer` | `McpServers.Teams.All` | ✅ Yes | ✅ Yes |
| `mcp_WordServer` | `McpServers.Word.All` | ✅ Yes | ✅ Yes |
| `mcp_ODSPRemoteServer` | `McpServers.OneDriveSharepoint.All` | ✅ Yes | ✅ Yes |
| `mcp_SharePointListsTools` | `McpServers.SharepointLists.All` | ✅ Yes | ✅ Yes |
| `mcp_MeServer` | `McpServers.Me.All` | ✅ Yes | ✅ Yes |
| `mcp_M365Copilot` | `McpServers.CopilotMCP.All` | ✅ Yes | ✅ Yes |
| `mcp_DASearch` | `McpServers.DASearch.All` | ❌ No | ⚠️ Unknown |
| `mcp_TeamsCanaryServer` | `McpServers.Teams.All` | ❌ No | ❌ Preview |
| `mcp_KnowledgeTools` | `McpServers.Knowledge.All` | ❌ No | ⚠️ Unknown |
| `mcp_Admin365_GraphTools` | `McpServers.Admin365Graph.All` | ❌ No | ⚠️ Unknown |
| `mcp_AdminTools` | `McpServers.M365Admin.All` | ❌ No | ⚠️ Unknown |

---

## How to Discover Tools in Undocumented Servers

Since these servers lack documentation, you can discover their available tools using the MCP Management Server:

### Option 1: Via Agent 365 CLI
```powershell
# Add the server first
a365 develop add-mcp-servers mcp_KnowledgeTools

# The CLI will show available tools when adding
```

### Option 2: Via MCP Management Server in VS Code
1. Connect to the MCP Management Server (see [MCP-SERVERS-REFERENCE.md](./MCP-SERVERS-REFERENCE.md))
2. Use the `GetTools` operation with the server name
3. Inspect the returned tool definitions

### Option 3: Inspect at Runtime
Add the server to your manifest and log the tools discovered during agent initialization:
```python
# In your agent code, after tool registration
logger.info(f"Registered tools: {[tool.name for tool in self.agent.tools]}")
```

---

## Getting Help

Since these servers are undocumented, consider these resources:

1. **Frontier Program Community** - If you're part of the Frontier preview, use the community channels
2. **Microsoft Support** - Open a support ticket for specific questions
3. **GitHub Issues** - Check [Agent365-Samples](https://github.com/microsoft/Agent365-Samples) for examples
4. **Trial and Error** - Add the server and inspect available tools at runtime

---

## Disclaimer

> ⚠️ **The information in this document is speculative** and based on naming conventions and patterns from documented MCP servers. The actual capabilities may differ significantly.
>
> These servers are part of the Agent 365 Frontier preview and may:
> - Change without notice
> - Be removed from the catalog
> - Have different behavior than expected
> - Require additional permissions not listed here
>
> Always test thoroughly in non-production environments first.
