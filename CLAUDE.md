# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Office MCP Server is a Model Context Protocol (MCP) server that provides Claude with access to Microsoft 365 services through the Microsoft Graph API. The server supports headless operation with automatic token refresh after initial authentication.

## Development Commands

### Build and Run
```bash
# Start server (stdio transport)
npm start

# Start auth server for initial OAuth flow
npm run auth-server

# Test mode with mock data
npm run test-mode

# MCP Inspector for testing
npm run inspect
```

### Testing
```bash
# Run all tests
npm test

# Run tests in watch mode
npm test:watch

# Run with coverage
npm test:coverage

# Run specific test suites
npm run test:auth
npm run test:teams
npm run test:drive
npm run test:planner
npm run test:users
npm run test:notifications
```

## Architecture

### Module-Based Design

The server is organized into isolated modules in `/[module]/index.js`, each exporting a `[module]Tools` array. Modules are:

- **auth**: OAuth 2.0 flow, token management with auto-refresh
- **email**: Unified email search with KQL support, attachment handling, SharePoint URL mapping
- **calendar**: Full CRUD operations, Teams meeting integration, recurrence patterns
- **teams**: Consolidated tools for meetings, channels, and chats
- **contacts**: Full contact management with advanced search
- **planner**: Microsoft Planner task management
- **files**: OneDrive/SharePoint operations with local path mapping
- **search**: Unified Microsoft 365 search capabilities
- **notifications**: Graph API change notifications/webhooks
- **utils**: Shared utilities (Graph API client, batch operations, mock data)

### Tool Registration Pattern

All modules follow this pattern:

```javascript
// /[module]/index.js
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

const [module]Tools = [
  {
    name: 'tool_name',
    description: 'Tool description',
    inputSchema: { /* JSON schema */ },
    handler: async (args) => {
      const accessToken = await ensureAuthenticated();
      // Implementation
    }
  }
];

module.exports = { [module]Tools };
```

Main entry point (`index.js`) imports all tools and registers them with the MCP Server SDK.

### Authentication Flow

1. **Initial auth**: User runs `office-auth-server.js` on port 3000, completes OAuth in browser
2. **Token storage**: Tokens saved to `~/.office-mcp-tokens.json` (configurable via `AUTH_CONFIG.tokenStorePath`)
3. **Auto-refresh**: `auth/auto-refresh.js` automatically refreshes tokens before expiry
4. **Headless operation**: After initial auth, server runs without browser interaction

All tools use `ensureAuthenticated()` which transparently handles token refresh.

### Graph API Client

`utils/graph-api.js` provides `callGraphAPI()` with:
- Automatic retry with exponential backoff for rate limits (429) and service errors (503, 504)
- Enhanced error messages with actionable suggestions
- Test mode support routing to `utils/mock-data.js`
- Proper OData filter encoding for Graph API

### Local Path Mapping

Email and file modules map SharePoint/OneDrive URLs to local sync paths:
- `SHAREPOINT_SYNC_PATH`: Local SharePoint sync folder
- `ONEDRIVE_SYNC_PATH`: Local OneDrive sync folder
- `TEMP_ATTACHMENTS_PATH`: Temp directory for email attachments

Pattern: `convertSharePointUrlToLocal()` in `email/index.js` handles URL parsing and path construction.

### Configuration

`config.js` centralizes all configuration:
- Environment variables with sensible defaults
- Graph API select field patterns for each module
- Retry/pagination settings
- Local path configuration

All modules import and reference `config` for consistency.

### Teams Module Structure

Teams has a consolidated structure (`teams/consolidated/`) with operation-based tools:
- `teams_meeting.js`: Complete meeting management (create, update, cancel, transcripts, recordings)
- `teams_channel.js`: Channel operations (messages, members, tabs)
- `teams_chat.js`: Chat management (create, send, members)

This reduces tool count from many individual tools to 3 consolidated, operation-based tools.

### Email Search Architecture

Email module uses a unified `email_search` tool that automatically routes queries:
- **KQL route**: Advanced queries with operators → `/me/messages?$search="..."`
- **Filter route**: Date-filtered queries → `/me/messages?$filter=receivedDateTime ge ...`
- **Smart fallback**: Switches routes based on query patterns and date requirements

## Testing Strategy

### Test Mode
Set `USE_TEST_MODE=true` to use mock data from `utils/mock-data.js` without real API calls. All Graph API calls route to `simulateGraphAPIResponse()`.

### Test Structure
- Tests in `/tests/*.test.js`
- Use `tests/test-utils.js` for shared test helpers
- Mock token manager and Graph API in unit tests
- Integration tests verify end-to-end tool execution

### Running Specific Tests
Each module has a dedicated test script (`npm run test:[module]`) for focused testing during development.

## Important Patterns

### Error Handling
Always wrap Graph API calls in try-catch and provide user-friendly error messages. The Graph API client already enhances errors, but tools should add context:

```javascript
try {
  const result = await callGraphAPI(accessToken, 'GET', path);
  return { content: [{ type: "text", text: formatResult(result) }] };
} catch (error) {
  return {
    content: [{ type: "text", text: `Failed to fetch data: ${error.message}` }],
    isError: true
  };
}
```

### Batch Operations
Use `utils/batch.js` for operations on multiple items (e.g., deleting multiple emails, updating multiple tasks). Batch requests reduce API calls and improve performance.

### Date Handling
Graph API requires ISO 8601 format. Email and calendar modules support both:
- Absolute: `YYYY-MM-DD` (converted to ISO 8601)
- Relative: `7d`, `1w`, `1m`, `1y` (converted to dates, then ISO 8601)

### Tool Response Format
All tools return MCP-compliant responses:
```javascript
{
  content: [
    { type: "text", text: "Result message" }
  ],
  isError: false  // Set to true for errors
}
```

## Common Development Tasks

### Adding a New Tool
1. Add tool definition to appropriate module's tools array
2. Implement handler following the pattern above
3. Add tests in `/tests/[module].test.js`
4. Update tool count validation in tests if needed

### Adding a New Module
1. Create `/[module]/index.js` with tools array export
2. Import in main `index.js` and add to TOOLS array
3. Create test file `/tests/[module].test.js`
4. Add test script to `package.json`

### Debugging Authentication
- Check token file exists: `~/.office-mcp-tokens.json`
- Verify Azure app permissions match `config.js` scopes
- Use `npm run auth-server` to re-authenticate
- Check `office-auth-server.js` logs for OAuth callback errors

### Adding SharePoint Path Mapping
The `convertSharePointUrlToLocal()` function handles both organizational SharePoint sites and personal OneDrive. When adding new path patterns:
1. Test with real URLs from target environment
2. Handle URL decoding for special characters
3. Consider both Windows and Unix path formats
