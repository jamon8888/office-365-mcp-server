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
npm run test:contacts
npm run test:newsletter
```

## Architecture

### Module-Based Design

The server is organized into isolated modules in `/[module]/index.js`, each exporting a `[module]Tools` array. Modules are:

- **auth**: OAuth 2.0 flow, token management with auto-refresh (3 tools)
- **email**: Unified email operations with KQL search, attachment handling, SharePoint URL mapping, contact extraction with newsletter filtering, categories, focused inbox, draft management (8 tools)
- **calendar**: Full CRUD operations, Teams meeting integration, recurrence patterns (1 consolidated tool)
- **teams**: Consolidated tools for meetings, channels, and chats (3 tools)
- **contacts**: Full contact management with advanced search (1 consolidated tool)
- **planner**: Microsoft Planner task management with enhanced assignments and bulk operations (8 tools)
- **files**: OneDrive/SharePoint operations with local path mapping and symlink support (2 tools)
- **search**: Unified Microsoft 365 search with aggregations, enrichment, and multi-entity support (1 tool)
- **notifications**: Graph API change notifications/webhooks (subscription management tools)
- **utils**: Shared utilities (Graph API client, batch operations, mock data, contact extraction, newsletter detection)

**Total: ~28 tools across 9 modules**

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

### Module Tools Reference

Complete list of all tools available in each module:

**Auth Module** (3 tools):
- `about` - Server information and version
- `authenticate` - Microsoft Graph API authentication
- `check-auth-status` - Check current authentication status

**Email Module** (8 tools):
- `email` - Unified email tool with operations: list, read, send, draft, update_draft, send_draft, list_drafts
- `email_search` - Advanced search with KQL support, folder filtering, and automatic optimization
- `email_move` - Move emails between folders with batch processing support
- `email_folder` - Manage folders (operations: list, create)
- `email_rules` - Manage inbox rules (operations: list, create, with enhanced mode)
- `email_focused` - Get AI-filtered focused inbox messages
- `email_categories` - Manage categories (operations: list, create, update, delete, apply, remove)
- `extract_contacts_from_emails` - Extract contact information with intelligent newsletter filtering

**Calendar Module** (1 tool):
- `calendar` - Unified calendar tool with operations: list, create, get, update, delete

**Teams Module** (3 tools):
- `teams_meeting` - Operations: create, update, cancel, get, find_by_url, list_transcripts, get_transcript, list_recordings, get_recording, get_participants, get_insights
- `teams_channel` - Operations: list, create, get, update, delete, list_messages, get_message, create_message, reply_to_message, list_members, add_member, remove_member, list_tabs
- `teams_chat` - Operations: list, create, get, update, delete, list_messages, get_message, send_message, update_message, delete_message, list_members, add_member, remove_member

**Contacts Module** (1 tool):
- `contacts` - Operations: list, search, get, create, update, delete, list_folders, create_folder

**Planner Module** (8 tools):
- `planner_plan` - Operations: list, get, create, update, delete
- `planner_task` - Operations: list, create, update, delete, get, assign
- `planner_bucket` - Operations: list, create, update, delete, get_tasks
- `planner_user` - User lookup (single or multiple)
- `planner_task_enhanced` - Operations: create, update_assignments (with removeAssignments support)
- `planner_assignments` - Operations: get, update
- `planner_task_details` - Get detailed task information including checklist and references
- `planner_bulk_operations` - Operations: update, delete (for multiple tasks simultaneously)

**Files Module** (2 tools):
- `files` - Operations: list, get, upload, download, delete, share, search, move, copy, create_folder
- `files_map_sharepoint_path` - Map SharePoint URLs to local sync paths (supports Windows, WSL, and symlink formats)

**Search Module** (1 tool):
- `search` - Unified Microsoft 365 search with intelligent routing, aggregations, content enrichment, and support for: driveItem, listItem, message, event, person, chatMessage

**Notifications Module**:
- `create_subscription` - Create Graph API change notification webhooks
- `list_subscriptions` - List active notification subscriptions
- Additional subscription management tools

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
- `SHAREPOINT_SYNC_PATH`: Local SharePoint sync folder (e.g., `C:\Users\[user]\OneDrive - [Org]`)
- `ONEDRIVE_SYNC_PATH`: Local OneDrive sync folder (e.g., `C:\Users\[user]\OneDrive`)
- `SHAREPOINT_SYMLINK_PATH`: Optional symlink path for SharePoint (for WSL/Linux environments)
- `TEMP_ATTACHMENTS_PATH`: Temp directory for email attachments

The `files_map_sharepoint_path` tool supports multiple output formats:
- **Windows**: Native Windows path format
- **WSL**: Windows Subsystem for Linux path format (/mnt/c/...)
- **Symlink**: Custom symlink path for containerized environments

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

### Email Module Architecture

**Unified Email Tool**: The `email` tool consolidates common operations (list, read, send) and supports draft management:
- **Draft operations**: create drafts, update existing drafts, send drafts, list all drafts
- **Send operations**: direct send or send from draft
- **List/Read**: retrieve and display emails with full metadata

**Advanced Search**: The `email_search` tool provides intelligent query routing:
- **KQL route**: Advanced queries with operators → `/me/messages?$search="..."`
- **Filter route**: Date-filtered queries → `/me/messages?$filter=receivedDateTime ge ...`
- **Folder filtering**: Search within specific folders (Inbox, Sent, Archive, etc.)
- **Smart fallback**: Automatically switches routes based on query patterns and date requirements

**Focused Inbox**: The `email_focused` tool leverages Microsoft's AI to retrieve priority messages from the focused inbox, reducing noise from low-priority emails.

**Category Management**: The `email_categories` tool provides full CRUD operations for email categories:
- List all master categories
- Create custom categories with color coding
- Update category properties (name, color)
- Delete categories
- Apply/remove categories to/from emails (batch support)

### Contact Extraction and Newsletter Detection

The email module includes sophisticated contact extraction with intelligent newsletter filtering.

**Contact Extraction** (`tools/email-contact-extractor.js`):
- Extracts contacts from email metadata (From, To, CC) and body content
- Parses signatures, phone numbers, LinkedIn URLs, company names, job titles
- Supports bilingual extraction (French and English patterns)
- Cross-references with Outlook contacts to identify new contacts
- Deduplicates contacts with confidence scoring
- Exports to CSV with full contact information

**Utility Modules**:
- `utils/contact-parser.js`: Email, phone, LinkedIn extraction; name parsing; signature detection
- `utils/html-processor.js`: HTML stripping, signature extraction, body processing
- `utils/deduplicator.js`: Contact merging, deduplication, confidence scoring
- `utils/csv-generator.js`: CSV generation with proper escaping
- `utils/newsletter-detector.js`: Intelligent newsletter/marketing email detection

**Newsletter Detection** (`utils/newsletter-detector.js`):
- Multi-factor scoring system analyzing 14+ signals
- Bilingual pattern detection (French and English)
- Configurable via `/config/newsletter-rules.json` (whitelist, blacklist, custom patterns)
- Performance optimized with LRU caching (max 1000 entries)
- Default threshold: 60/100 (configurable)

**Detection Signals**:
- Headers: List-Unsubscribe, Precedence: bulk, ESP markers (Mailchimp, Sendgrid)
- Sender patterns: noreply@, newsletter@, ne-pas-repondre@, marketing@, bulletin@
- Body content: Unsubscribe links, "view in browser", preferences management
- Structure: High image/text ratio, tracking pixels, table-based layouts
- Recipients: BCC recipients, generic names ("Cher Client", "Valued Customer")

**Configuration** (`config/newsletter-rules.json`):
- Whitelist domains/senders (never filtered)
- Blacklist domains/senders (always filtered)
- Custom regex patterns for sender addresses and subject lines
- Configurable threshold and cache settings

**Integration Pattern**:
```javascript
const { detectNewsletter, applyWhitelistBlacklist } = require('../utils/newsletter-detector');

// Filter newsletters before processing
for (const email of emails) {
  const detection = await detectNewsletter(email, threshold);
  const finalDetection = applyWhitelistBlacklist(email, detection, rules);

  if (!finalDetection.isNewsletter) {
    // Process email for contact extraction
  }
}
```

### Unified Search Module

The `search` tool provides advanced Microsoft 365 search capabilities across multiple entity types:

**Supported Entity Types**:
- `driveItem`: Files and folders in OneDrive/SharePoint
- `listItem`: SharePoint list items
- `message`: Email messages
- `event`: Calendar events
- `person`: People and contacts
- `chatMessage`: Teams chat messages

**Advanced Features**:
- **Intelligent routing**: Automatically selects optimal search endpoint based on entity types
- **Aggregations**: Faceted search with dynamic filtering (fileType, lastModifiedBy, createdDateTime, etc.)
- **Content enrichment**: Extract Excel table data, document previews, and snippets
- **Query suggestions**: Get alternative query suggestions when results are limited
- **Deduplication**: Collapse duplicate results based on content similarity
- **Field selection**: Custom field selection for optimized responses

**Search Strategies**:
- Single entity type: Direct search with entity-specific optimizations
- Multiple entity types: Parallel search across all types with combined results
- Aggregation-enabled: Provides faceted search results for filtering

### Planner Module Enhancements

Beyond basic task management, the Planner module includes advanced features:

**Enhanced Assignments** (`planner_task_enhanced`):
- Create tasks with initial assignments
- Update assignments with `removeAssignments` parameter to remove specific assignees
- Supports partial assignment updates without affecting other assignees

**Task Details** (`planner_task_details`):
- Retrieve comprehensive task information including:
  - Full checklist items with completion status
  - Reference documents and links
  - Preview type and description
  - Complete task metadata

**Bulk Operations** (`planner_bulk_operations`):
- Update multiple tasks simultaneously (e.g., reassign multiple tasks, change due dates)
- Delete multiple tasks in a single operation
- Batch processing for improved performance
- Progress reporting during bulk operations

**Assignment Management** (`planner_assignments`):
- Get current assignments for a task
- Update assignments with fine-grained control
- Supports adding and removing assignees independently

## Testing Strategy

### Test Mode
Set `USE_TEST_MODE=true` to use mock data from `utils/mock-data.js` without real API calls. All Graph API calls route to `simulateGraphAPIResponse()`.

### Test Structure
- Tests in `/tests/*.test.js`
- Use `tests/test-utils.js` for shared test helpers
- Mock token manager and Graph API in unit tests
- Integration tests verify end-to-end tool execution

**Key Test Files**:
- `tests/contact-extraction.test.js`: Contact parsing, HTML processing, deduplication, CSV generation (42 tests)
- `tests/newsletter-detector.test.js`: Newsletter detection patterns, bilingual support, whitelist/blacklist (43 tests)
- Other module tests: auth, teams, drive, planner, users, notifications

### Running Specific Tests
Each module has a dedicated test script (`npm run test:[module]`) for focused testing during development.

### French Localization Testing
The contact extraction and newsletter detection modules include comprehensive French language support:
- French signature markers: Cordialement, Bien cordialement, Salutations, Amitiés
- French phone formats: +33, 01-09 (landline), 06-07 (mobile)
- French company types: SA, SARL, SAS, SASU, SNC
- French job titles: PDG, DG, Directeur, Responsable, Ingénieur
- French newsletter patterns: ne-pas-repondre@, désabonnement, bulletin@

All bilingual patterns are tested in both French and English.

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
4. Support symlink paths for containerized environments (use `SHAREPOINT_SYMLINK_PATH`)

The `files_map_sharepoint_path` tool provides a standardized way to convert SharePoint URLs to local paths with support for different environments (Windows, WSL, symlink).

### Extending Contact Extraction
To add new extraction patterns or enhance existing ones:
1. Update pattern regexes in `utils/contact-parser.js` for new formats
2. Add corresponding tests in `tests/contact-extraction.test.js`
3. For French patterns, ensure both accented and non-accented variants are matched
4. Update confidence scoring in `utils/deduplicator.js` if adding new high-value fields

### Customizing Newsletter Detection
To adjust newsletter detection behavior:
1. Modify signal weights in `NEWSLETTER_SIGNALS` object in `utils/newsletter-detector.js`
2. Add new detection patterns to appropriate check functions (checkHeaders, checkSenderPatterns, etc.)
3. Update `/config/newsletter-rules.json` for domain-specific whitelist/blacklist
4. Add tests for new patterns in `tests/newsletter-detector.test.js`
5. Adjust default threshold in config or per-call if needed

### Adding Bilingual Support for Other Languages
The contact extraction and newsletter detection modules use a bilingual pattern approach:
1. Add new language patterns alongside existing French/English patterns
2. Use non-capturing groups for language alternatives: `(?:English|French|NewLang)`
3. Test all language variants independently
4. Document supported languages in README.md and tool descriptions

### Working with Email Categories
The `email_categories` tool provides comprehensive category management:
- Categories support custom colors from predefined palette
- Apply/remove operations support batch processing for efficiency
- Master category list is user-specific (unique per mailbox)
- Color values: preset_0 through preset_24 (25 total color options)

### Implementing Bulk Operations
When working with planner bulk operations:
1. Use `planner_bulk_operations` for multiple task updates/deletes
2. Monitor progress through incremental reporting
3. Handle partial failures gracefully (some tasks may fail while others succeed)
4. Always validate task IDs and eTags before bulk operations
5. Consider rate limits when operating on large task sets

### Leveraging Search Aggregations
The search tool's aggregation feature enables faceted search:
1. Request aggregations by specifying bucket fields (e.g., fileType, lastModifiedBy)
2. Use returned aggregations to build dynamic filters
3. Combine aggregations with custom field selection for optimized queries
4. Enable content enrichment for rich result previews

## Additional Documentation

The repository includes additional documentation files:
- **SETUP.md**: Detailed setup instructions for configuring the server
- **CONTACTS_API.md**: Comprehensive API documentation for contact operations
- **PERFORMANCE.md**: Performance optimization guidelines and benchmarks
- **OPTIMIZATION_SUMMARY.md**: Summary of implemented performance optimizations
- **planner/ASSIGNMENT_FIX_SUMMARY.md**: Technical details on planner assignment handling fixes
