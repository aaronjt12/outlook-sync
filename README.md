# SWA Ticket Sync

A production-ready React application that synchronizes unread emails from Microsoft Outlook with SharePoint lists, creating support tickets automatically. Built for IT support teams to streamline ticket creation from incoming emails.

## Overview

SWA Ticket Sync bridges the gap between email-based support requests and SharePoint ticketing systems. It features dual-account authentication, allowing users to sign into Outlook and SharePoint independently, select any site and list, configure field mappings, and batch-create tickets from unread emails.

## Features

### Core Functionality

- **Dual-Account Authentication**: Sign into Outlook and SharePoint separately with different accounts
- **Dynamic Site/List Selection**: Browse and select any SharePoint site and list
- **Unread Email Filtering**: Loads only unread emails from inbox, sorted oldest-first
- **Smart Field Mapping**: Auto-detects standard columns or allows custom mapping
- **Batch Ticket Creation**: Create multiple SharePoint tickets from selected emails
- **Auto-Mark as Read**: Automatically marks synced emails as read
- **Route Extraction**: Parses route information from email body
- **Persistent Configuration**: Saves site/list selection and field mappings per list
- **Toast Notifications**: Real-time feedback for all operations

### User Workflow

1. **Sign in to Outlook**: Authenticate with any Microsoft 365 account
2. **Sign in to SharePoint**: Authenticate with same or different account
3. **Select Site**: Choose from all available SharePoint sites
4. **Select List**: Choose target list from the selected site
5. **Configure Mapping**: Map email fields to SharePoint columns (auto-maps if possible)
6. **Load Unread Emails**: Fetch unread messages from inbox
7. **Select Emails**: Choose which emails to convert to tickets
8. **Sync**: Create tickets and mark emails as read

## Technical Architecture

### Technology Stack

- **Frontend Framework**: React 18.3.1
- **Build Tool**: Create React App (react-scripts 5.0.1)
- **Authentication**:
  - @azure/msal-browser 4.14.0
  - @azure/msal-react 3.0.14
- **API Integration**: Microsoft Graph API (REST)
- **Language**: JavaScript (ES6+)
- **State Management**: React Hooks (useState, useEffect)
- **Storage**: localStorage for persistence

### Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                   SWA Ticket Sync App                        │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌────────────────────────────────────────────────┐         │
│  │         Dual Authentication Panels             │         │
│  ├─────────────────┬──────────────────────────────┤         │
│  │  Outlook Auth   │   SharePoint Auth            │         │
│  │  (Mail.Read)    │   (Sites.ReadWrite.All)      │         │
│  └────────┬────────┴──────────────┬───────────────┘         │
│           │                       │                          │
│           │                       │                          │
│  ┌────────▼───────────────────────▼───────────┐            │
│  │         Main App Component                  │            │
│  │  • Dual-token management                    │            │
│  │  • Site/List selection                      │            │
│  │  • Field mapping orchestration              │            │
│  │  • Email sync coordination                  │            │
│  └────────┬────────────────────────────────────┘            │
│           │                                                  │
│           ├──────────────┬────────────────┬─────────────┐   │
│           ▼              ▼                ▼             ▼   │
│  ┌────────────┐  ┌──────────────┐  ┌─────────┐  ┌────────┐│
│  │ EmailTable │  │ MappingPanel │  │  Toast  │  │ Graph  ││
│  │ Component  │  │  Component   │  │Component│  │Helpers ││
│  └────────────┘  └──────────────┘  └─────────┘  └────────┘│
│                                                               │
└───────────────────────┬───────────────────────────────────┘
                        │
                        ▼
        ┌───────────────────────────────────────┐
        │      Microsoft Graph API v1.0          │
        ├───────────────────────────────────────┤
        │  Outlook:                              │
        │    • /me/mailFolders/inbox/messages    │
        │    • /me/messages/{id} (PATCH)         │
        │  SharePoint:                           │
        │    • /sites?search=*                   │
        │    • /sites/{id}/lists                 │
        │    • /sites/{id}/lists/{id}/columns    │
        │    • /sites/{id}/lists/{id}/items      │
        └───────────────────────────────────────┘
```

### Component Architecture

#### Core Components

**1. App Component** ([src/App.js](swa-ticket-sync/src/App.js))
- **Responsibilities**:
  - Dual authentication flow (Outlook + SharePoint)
  - Site and list selection with dynamic loading
  - Field mapping management
  - Email syncing orchestration
  - Toast notification management
- **Key State**:
  - `outlookToken` / `sharepointToken`: Separate auth tokens
  - `sites`, `lists`, `columns`: SharePoint data
  - `emails`, `selectedEmailIds`: Email data and selection
  - `fieldMapping`: Column mapping configuration
  - `toasts`: Notification queue

**2. EmailTable Component** ([src/components/EmailTable.jsx](swa-ticket-sync/src/components/EmailTable.jsx))
- **Responsibilities**: Display unread emails in a selectable table
- **Features**:
  - Checkbox selection (individual + select all)
  - Shows: Subject, Sender, Received date, Route preview
  - Visual selection highlighting
  - Empty state handling

**3. MappingPanel Component** ([src/components/MappingPanel.jsx](swa-ticket-sync/src/components/MappingPanel.jsx))
- **Responsibilities**: Configure field mappings between email and SharePoint
- **Features**:
  - Auto-mapping on first use
  - Dropdown selects for each field
  - Required field validation
  - Modal interface
  - Save per (siteId, listId)

**4. Toast Component** ([src/components/Toast.jsx](swa-ticket-sync/src/components/Toast.jsx))
- **Responsibilities**: User notifications
- **Types**: success, error, warning, info
- **Features**: Auto-dismiss, manual close, stacking

#### Library Modules

**1. Graph Helpers** ([src/lib/graph.js](swa-ticket-sync/src/lib/graph.js))
- `getSites(token)`: List SharePoint sites
- `getLists(siteId, token)`: List lists in a site
- `getListColumns(siteId, listId, token)`: Get list columns
- `getUnreadEmails(token, top)`: Fetch unread inbox messages
- `createListItem(siteId, listId, fields, token)`: Create SharePoint item
- `markEmailRead(messageId, token)`: Mark email as read
- `generateTicketNumber(date)`: Generate YYYYMMDDHHmm ticket number
- `extractRoute(bodyPreview)`: Extract route via regex

**2. Mapping Store** ([src/lib/mappingStore.js](swa-ticket-sync/src/lib/mappingStore.js))
- `saveMapping(siteId, listId, mapping)`: Persist field mapping
- `loadMapping(siteId, listId)`: Load saved mapping
- `saveLastSite(siteId, name)`: Save last selected site
- `saveLastList(listId, name)`: Save last selected list
- `clearAllMappings()`: Clear all localStorage data

### Authentication Flow

```
User loads app
     │
     ▼
┌─────────────────────────────────────────┐
│  Two separate auth panels displayed:     │
│  • Sign in to Outlook                    │
│  • Sign in to SharePoint                 │
└─────────────────────────────────────────┘
     │
     ▼
User clicks "Sign in to Outlook"
     │
     ▼
┌─────────────────────────────────────────┐
│  loginPopup({                            │
│    scopes: ['Mail.Read'],                │
│    prompt: 'select_account'              │
│  })                                      │
└─────────────────────────────────────────┘
     │
     ▼
Microsoft account picker → User selects account
     │
     ▼
outlookToken acquired → Shows "Connected" badge
     │
     ▼
User clicks "Sign in to SharePoint"
     │
     ▼
┌─────────────────────────────────────────┐
│  loginPopup({                            │
│    scopes: ['Sites.ReadWrite.All'],      │
│    prompt: 'select_account'              │
│  })                                      │
└─────────────────────────────────────────┘
     │
     ▼
Microsoft account picker → User selects account (can be different)
     │
     ▼
sharepointToken acquired → Auto-loads sites → Shows "Connected" badge
```

### Data Flow

#### Email to Ticket Sync Flow

```
User clicks "Sync X Selected Emails"
         │
         ▼
Validate: tokens, site, list, mapping
         │
         ▼
For each selected email:
    │
    ├─► Get email data
    │
    ├─► Generate ticket number: YYYYMMDDHHmm from receivedDateTime
    │
    ├─► Extract route from body: /route:\s*(\S+)/i
    │
    ├─► Build fields object using mapping:
    │     • ticketnumber → mapped column
    │     • subject → mapped column
    │     • route → mapped column (if found)
    │     • description → bodyPreview
    │     • user → sender email
    │     • status → "New"
    │
    ├─► POST /sites/{siteId}/lists/{listId}/items
    │     { fields: { ... } }
    │
    ├─► On success:
    │     └─► PATCH /me/messages/{id}
    │         { isRead: true }
    │
    └─► Track success/failure count
         │
         ▼
Show toast: "Created X tickets. (Y failed)"
         │
         ▼
Reload unread emails (synced emails no longer appear)
```

### API Integration

#### Microsoft Graph Endpoints

**Authentication Scopes:**
- Outlook: `https://graph.microsoft.com/Mail.Read`
- SharePoint: `https://graph.microsoft.com/Sites.ReadWrite.All`

**Endpoints Used:**

1. **Get SharePoint Sites**
   ```
   GET https://graph.microsoft.com/v1.0/sites?search=*
   ```

2. **Get Lists in Site**
   ```
   GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists
   ```

3. **Get List Columns**
   ```
   GET https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/columns
   ```

4. **Get Unread Emails**
   ```
   GET https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages
   Query: $filter=isRead eq false&$top=20&$orderby=receivedDateTime asc
   ```

5. **Create SharePoint List Item**
   ```
   POST https://graph.microsoft.com/v1.0/sites/{siteId}/lists/{listId}/items
   Body: { fields: { ticketnumber: "...", subject: "...", ... } }
   ```

6. **Mark Email as Read**
   ```
   PATCH https://graph.microsoft.com/v1.0/me/messages/{messageId}
   Body: { isRead: true }
   ```

### File Structure

```
swa-ticket-sync/
├── public/
│   └── index.html              # HTML template
├── src/
│   ├── components/
│   │   ├── EmailTable.jsx      # Email grid with selection
│   │   ├── EmailTable.css
│   │   ├── MappingPanel.jsx    # Field mapping modal
│   │   ├── MappingPanel.css
│   │   ├── Toast.jsx           # Notification system
│   │   └── Toast.css
│   ├── lib/
│   │   ├── graph.js            # Microsoft Graph API wrappers
│   │   └── mappingStore.js     # localStorage persistence
│   ├── index.js                # Entry point with MsalProvider
│   ├── index.css               # Global styles
│   ├── App.js                  # Main application component
│   ├── App.css                 # Application styles
│   └── authConfig.js           # MSAL configuration
├── package.json                # Dependencies
└── AZURE_SETUP.md              # Azure AD setup guide
```

### Key Design Decisions

1. **Dual-Account Authentication**: Uses `loginPopup` with `prompt: "select_account"` to allow different accounts for Outlook vs SharePoint
2. **Unread-Only Filtering**: Focuses on new emails only, oldest-first for queue workflow
3. **Auto-Mapping**: Reduces configuration time by detecting standard column names
4. **LocalStorage Persistence**: Saves site/list selection and mappings per list
5. **Component Separation**: Reusable components for table, mapping, and notifications
6. **Toast Notifications**: All user feedback surfaced in UI (no console-only errors)
7. **Validation Gates**: Prevents incomplete syncs with clear user guidance
8. **Route Extraction**: Regex parsing for optional route field from email body
9. **Military Time Format**: YYYYMMDDHHmm ticket numbering from email timestamp

## Setup Instructions

### Prerequisites

- Node.js 14+ and npm
- Azure AD application registration
- Microsoft 365 account(s) with:
  - Outlook mailbox access
  - SharePoint site access
- SharePoint list for tickets

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/aaronjt12/outlook-sync.git
   cd outlook-sync/swa-ticket-sync
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

3. **Configure Azure AD:**
   - Follow [AZURE_SETUP.md](swa-ticket-sync/AZURE_SETUP.md) for detailed instructions
   - Update `src/authConfig.js` with your Azure AD Client ID:

   ```javascript
   export const msalConfig = {
     auth: {
       clientId: "YOUR_CLIENT_ID_HERE",
       authority: "https://login.microsoftonline.com/common",
       redirectUri: window.location.origin,
     },
     cache: {
       cacheLocation: "sessionStorage",
       storeAuthStateInCookie: false,
     }
   };
   ```

4. **Start the development server:**
   ```bash
   npm start
   ```

5. **Open [http://localhost:3000](http://localhost:3000)**

### Azure AD Configuration

#### Required API Permissions

In your Azure AD app registration, add these Microsoft Graph **Delegated** permissions:

- `Mail.Read` - Read user email
- `Sites.ReadWrite.All` - Read and write items in all site collections

#### Redirect URI

Add your application URL as a redirect URI:
- Development: `http://localhost:3000`
- Production: `https://your-domain.com`

Platform type: **Single-page application (SPA)**

### SharePoint List Setup

Your SharePoint list can have any columns. The app will:
1. Load all available columns
2. Auto-map if it finds columns with internal names: `ticketnumber`, `subject`, `route`, `description`, `user`
3. Allow manual mapping to any columns if auto-mapping doesn't work

**Recommended columns:**
- **TicketNumber** (Single line of text)
- **Subject** (Single line of text)
- **Route** (Single line of text)
- **Description** (Multiple lines of text)
- **User** (Single line of text or Person field)
- **Status** (Choice: New, In Progress, Resolved)

## Usage Guide

### Creating Tickets from Emails

1. **Authenticate**
   - Click "Sign in to Outlook" and select your email account
   - Click "Sign in to SharePoint" and select your SharePoint account (can be different)

2. **Select Target**
   - Choose a SharePoint Site from the dropdown
   - Choose a List from the site
   - You'll see a mapping status indicator

3. **Configure Mapping** (first time only)
   - Click "Configure Mapping"
   - Map required fields: Ticket Number, Subject, Description, User
   - Optionally map: Route, Status
   - Click "Save Mapping" (persists for this list)

4. **Load Emails**
   - Click "Load Unread Emails"
   - Review the unread emails (oldest first)
   - Route information will be extracted if present in email body

5. **Select and Sync**
   - Check boxes next to emails you want to convert
   - Click "Sync X Selected Emails"
   - Wait for completion toast
   - Synced emails are marked as read and disappear from the list

### Field Mapping

The app maps email data to SharePoint columns:

| App Field | Source | Format |
|-----------|--------|--------|
| Ticket Number | Email received date | YYYYMMDDHHmm |
| Subject | Email subject line | Text |
| Route | Email body preview | Regex: `/route:\s*(\S+)/i` |
| Description | Email body preview | Plain text |
| User | Email sender | Email address |
| Status | Static | "New" |

### Route Extraction

If your emails contain route information in the format `route: ROUTE123` or `Route: ABC-456`, the app will automatically extract it.

Example email body:
```
We have an issue with the printer.
Route: BLDG-A-2F
Please help!
```
Extracted route: `BLDG-A-2F`

## Development

### Available Scripts

- **`npm start`**: Run development server (localhost:3000)
- **`npm test`**: Run test suite
- **`npm run build`**: Create production build
- **`npm run eject`**: Eject from Create React App (irreversible)

### Project Structure

- **Components**: Reusable UI components with isolated CSS
- **Lib**: Utility functions and API wrappers
- **App.js**: Main application logic and state management
- **authConfig.js**: MSAL authentication configuration

### Code Style

- ES6+ JavaScript with modern React patterns
- Functional components with React Hooks
- CSS Modules for component-specific styling
- Async/await for API calls
- Comprehensive error handling

## Security Considerations

- **OAuth 2.0 Authentication**: Via Microsoft Identity Platform
- **Token Storage**: sessionStorage (cleared on browser close)
- **Scoped Permissions**: Least privilege (Mail.Read, Sites.ReadWrite.All)
- **No Client Secret**: Public client application (SPA)
- **Account Picker**: Always prompts for account selection
- **HTTPS Required**: For production deployment
- **No Token Logging**: Tokens never logged to console

## Troubleshooting

### Common Issues

**"Outlook login failed"**
- Ensure Mail.Read permission is granted in Azure AD
- Check that your account has mailbox access
- Try clearing browser cache and sessionStorage

**"SharePoint login failed"**
- Ensure Sites.ReadWrite.All permission is granted
- Admin consent may be required for Sites scope
- Verify account has access to SharePoint

**"Failed to load sites"**
- Check SharePoint token is valid
- Ensure account has permission to list sites
- Verify network connectivity to graph.microsoft.com

**"No columns available"**
- List may have no editable columns
- System columns (like ContentType) are filtered out
- Try a different list with custom columns

**"Mapping validation fails"**
- Ensure all required fields are mapped
- Required: Ticket Number, Subject, Description, User
- Check that target columns exist and are writable

**"Ticket creation failed"**
- Verify field mapping is correct
- Check SharePoint column types match data
- Ensure account has write permission to list
- Review browser console for detailed errors

**"Emails not marking as read"**
- Mail.ReadWrite scope may be needed
- Current scope (Mail.Read) may not allow modifications
- Update authConfig.js scopes to include `Mail.ReadWrite`

## Browser Compatibility

- **Recommended**: Chrome 90+, Edge 90+, Firefox 88+, Safari 14+
- **Requires**: ES6 support, localStorage, sessionStorage
- **Mobile**: Responsive design supports iOS Safari and Chrome Android

## Performance

- **Initial Load**: ~115 KB gzipped JavaScript + 2 KB CSS
- **Email Fetch**: ~1-2 seconds for 20 emails
- **Site Load**: ~2-3 seconds (depends on tenant size)
- **Ticket Creation**: ~1-2 seconds per email (sequential)
- **Caching**: Sites, lists, columns cached in memory during session

## Contributing

This is a private project. For questions or support, contact the repository owner.

## License

ISC

## Project Links

- **Repository**: https://github.com/aaronjt12/outlook-sync
- **Issues**: https://github.com/aaronjt12/outlook-sync/issues

## Changelog

### v2.0 (2025) - Production Refactor
- **Breaking Changes**:
  - Removed technician assignment feature
  - Removed folder creation functionality
  - Changed from single-auth to dual-auth model
- **New Features**:
  - Dual-account authentication (Outlook + SharePoint separate)
  - Dynamic site and list selection
  - Smart field mapping with auto-detection
  - Unread-only email filtering
  - Auto-mark emails as read after sync
  - Route extraction from email body
  - Toast notification system
  - LocalStorage persistence for mappings
- **Architecture**:
  - Refactored to component-based architecture
  - Created graph.js API wrapper library
  - Created mappingStore.js persistence layer
  - Modern card-based UI with responsive design
- **Technical**:
  - Uses loginPopup instead of loginRedirect
  - Separate token management per service
  - Improved error handling and user feedback

### v1.0 (2025) - MVP
- Initial MSAL authentication
- Basic email fetching
- SharePoint ticket creation
- Field mapping functionality
- Technician CSV upload and assignment

## Support

For help with:
- **Azure AD Setup**: See [AZURE_SETUP.md](swa-ticket-sync/AZURE_SETUP.md)
- **MSAL Issues**: Check [@azure/msal-react documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-react)
- **Graph API**: See [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)
- **Bugs**: Open an issue on GitHub
