# SWA Ticket Sync

A React-based web application that synchronizes emails from Microsoft Outlook with SharePoint lists, creating support tickets automatically. Built for IT support teams to streamline ticket creation from incoming emails.

## Overview

SWA Ticket Sync bridges the gap between email-based support requests and SharePoint ticketing systems. It allows users to authenticate with their Microsoft account, fetch emails from Outlook, map email fields to SharePoint columns, assign tickets to technicians, and batch-create tickets in SharePoint.

## Features

### Core Functionality

- **Microsoft Account Authentication**: Secure OAuth 2.0 authentication via MSAL (Microsoft Authentication Library)
- **Email Fetching**: Retrieve up to 50 most recent emails from Outlook inbox
- **Email Selection**: Interactive checkbox interface to select multiple emails for ticket creation
- **Field Mapping**: Dynamic mapping between email properties and SharePoint list columns
- **Technician Management**: CSV-based technician roster with assignment capabilities
- **Batch Ticket Creation**: Create multiple SharePoint tickets from selected emails in one operation
- **Auto-Generated Ticket Numbers**: Military time format ticket numbering (YYYYMMDDHHMM)

### User Workflow

1. **Login**: Authenticate with Microsoft account (account picker enabled)
2. **Fetch Emails**: Retrieve recent emails from Outlook
3. **Select Emails**: Choose which emails should become tickets
4. **Upload Technicians**: Load technician roster from CSV file
5. **Map Fields**: Configure how email data maps to SharePoint columns
6. **Assign Technicians**: Assign tickets to specific technicians
7. **Create Tickets**: Batch create tickets in SharePoint

## Technical Architecture

### Technology Stack

- **Frontend Framework**: React 18.3.1
- **Build Tool**: Create React App (react-scripts 5.0.1)
- **Authentication**:
  - @azure/msal-browser 4.14.0
  - @azure/msal-react 3.0.14
- **API Integration**: Microsoft Graph API (REST)
- **Language**: JavaScript (ES6+)

### Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                     SWA Ticket Sync App                      │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌───────────────┐         ┌──────────────────────────┐    │
│  │   MSAL Auth   │────────▶│   Microsoft Identity     │    │
│  │   Provider    │         │   Platform (Azure AD)    │    │
│  └───────────────┘         └──────────────────────────┘    │
│         │                                                    │
│         ▼                                                    │
│  ┌───────────────────────────────────────────────────┐     │
│  │              Main App Component                    │     │
│  │  - Authentication State Management                 │     │
│  │  - Token Acquisition & Refresh                     │     │
│  │  - Component Orchestration                         │     │
│  └───────────────────────────────────────────────────┘     │
│         │                                                    │
│         ├──────────────┬──────────────┬──────────────┐     │
│         ▼              ▼              ▼              ▼      │
│  ┌──────────┐  ┌─────────────┐  ┌──────────┐  ┌────────┐  │
│  │  Email   │  │ Technician  │  │  Field   │  │ Create │  │
│  │  List    │  │ Assignment  │  │ Mapping  │  │ Button │  │
│  │Component │  │  Component  │  │Component │  │Component│  │
│  └──────────┘  └─────────────┘  └──────────┘  └────────┘  │
│                                                               │
└───────────────────────┬───────────────────────────────────┘
                        │
                        ▼
        ┌───────────────────────────────────┐
        │    Microsoft Graph API             │
        ├───────────────────────────────────┤
        │  • /me/messages (Outlook)          │
        │  • /sites/root/lists/Tickets/...   │
        │    - columns (SharePoint)          │
        │    - items (SharePoint)            │
        └───────────────────────────────────┘
```

### Component Architecture

#### Core Components

**1. App Component** (`src/App.js`)
- **Responsibilities**:
  - Authentication flow orchestration
  - State management for entire application
  - API calls to Microsoft Graph
  - Component composition and layout
- **Key State**:
  - `accessToken`: OAuth access token for API calls
  - `emails`: Array of fetched Outlook messages
  - `selectedEmails`: User-selected email IDs
  - `sharePointFields`: Available SharePoint columns
  - `fieldMapping`: Email-to-SharePoint field mapping configuration
  - `technicians`: Loaded technician roster
  - `emailTechnicianMapping`: Ticket assignment configuration
- **Hooks Used**:
  - `useMsal`: Access MSAL instance and authenticated accounts
  - `useState`: Local component state
  - `useEffect`: Token acquisition side effects

**2. TechnicianAssignment Component** (inline in App.js)
- **Responsibilities**:
  - CSV file upload and parsing
  - Technician selection management
  - Priority ordering (up/down arrows)
  - Per-email technician assignment
- **Features**:
  - Drag-free priority ordering
  - Multi-select with checkboxes
  - Visual indication of selected technicians

**3. FieldMapping Component** (inline in App.js)
- **Responsibilities**:
  - Display available app fields
  - Display fetched SharePoint columns
  - Configure field-to-column mappings
- **Mapped Fields**:
  - `subject`: Email subject line
  - `description`: Email body content
  - `user`: Email sender address
  - `ticketnumber`: Auto-generated timestamp
  - `assigned to`: Selected technician name

**4. CreateTicketsButton Component** (inline in App.js)
- **Responsibilities**:
  - Validate prerequisites (token, selections, mappings)
  - Batch ticket creation
  - Error handling and user feedback
  - Loading state management

### Authentication Flow

```
┌────────────┐
│  App Load  │
└─────┬──────┘
      │
      ▼
┌─────────────────────────┐
│ MsalProvider Initializes│
│ (index.js)               │
└─────┬───────────────────┘
      │
      ▼
┌────────────────────────┐      No      ┌──────────────────┐
│ User Authenticated?    │─────────────▶│ Show Login Button│
└────────┬───────────────┘              └──────────────────┘
         │ Yes                                     │
         │                                         │ Click
         │                                         ▼
         │                           ┌──────────────────────────┐
         │                           │ loginRedirect()          │
         │                           │ - prompt: select_account │
         │                           └──────────┬───────────────┘
         │                                      │
         │                                      ▼
         │                           ┌────────────────────────┐
         │                           │ Microsoft Login Page   │
         │                           │ (Azure AD)             │
         │                           └──────────┬─────────────┘
         │                                      │ Success
         │                                      ▼
         │                           ┌────────────────────────┐
         │              ┌────────────│ Redirect to App        │
         │              │            └────────────────────────┘
         │              │
         ▼              ▼
┌─────────────────────────────────┐
│ acquireTokenSilent()             │
│ - Gets access token from cache  │
│ - Or silently refreshes token   │
└─────────┬───────────────────────┘
          │
          ▼
┌─────────────────────────────────┐
│ setAccessToken()                 │
│ - App renders authenticated UI  │
└─────────────────────────────────┘
```

### Data Flow

#### Email Fetching Flow
```
User Click "Fetch Emails"
         ↓
GET /v1.0/me/messages?$top=50&$orderby=receivedDateTime desc
         ↓
Response: { value: [...emails] }
         ↓
setEmails(data.value)
         ↓
Render Email List Component
```

#### Ticket Creation Flow
```
User Click "Create Tickets"
         ↓
For each selectedEmail:
  1. Find email data
  2. Generate ticket number (YYYYMMDDHHMM)
  3. Get assigned technician
  4. Build payload using fieldMapping
  5. POST /v1.0/sites/root/lists/Tickets/items
         ↓
All requests complete
         ↓
Show success/error alert
```

### API Integration

#### Microsoft Graph Endpoints

**Authentication Scopes Required:**
- `https://graph.microsoft.com/Mail.Read`
- `https://graph.microsoft.com/Sites.ReadWrite.All`

**Endpoints Used:**

1. **Get Messages** (Outlook)
   ```
   GET https://graph.microsoft.com/v1.0/me/messages
   Query: $top=50&$orderby=receivedDateTime desc
   ```

2. **Get SharePoint Columns**
   ```
   GET https://graph.microsoft.com/v1.0/sites/root/lists/Tickets/columns
   ```

3. **Create SharePoint Item**
   ```
   POST https://graph.microsoft.com/v1.0/sites/root/lists/Tickets/items
   Body: { fields: { ...mappedData } }
   ```

### File Structure

```
swa-ticket-sync/
├── public/
│   └── index.html              # HTML template (cleaned, no CDN scripts)
├── src/
│   ├── index.js                # Entry point, MsalProvider wrapper
│   ├── index.css               # Global styles
│   ├── App.js                  # Main application component
│   ├── App.css                 # Application styles
│   └── authConfig.js           # MSAL configuration
├── package.json                # Dependencies and scripts
└── AZURE_SETUP.md              # Azure AD setup instructions
```

### Key Design Decisions

1. **MSAL React Integration**: Uses `@azure/msal-react` hooks rather than vanilla MSAL for better React integration
2. **Redirect Flow**: Uses `loginRedirect` instead of popup for better mobile compatibility
3. **Silent Token Acquisition**: Automatically refreshes tokens without user interaction
4. **Component Composition**: Main functionality organized as inline components for simplicity
5. **Military Time Ticketing**: Ticket numbers use YYYYMMDDHHMM format from email received time
6. **Session Storage**: MSAL cache uses sessionStorage for security (cleared on browser close)

## Setup Instructions

### Prerequisites

- Node.js 14+ and npm
- Azure AD application registration
- Microsoft 365 account with Outlook access
- SharePoint site with a "Tickets" list

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/aaronjt12/outlook-sync.git
   cd outlook-sync/swa-ticket-sync
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Configure Azure AD:
   - Follow [AZURE_SETUP.md](swa-ticket-sync/AZURE_SETUP.md) for detailed instructions
   - Update `src/authConfig.js` with your Azure AD Client ID

4. Start the development server:
   ```bash
   npm start
   ```

5. Open [http://localhost:3000](http://localhost:3000)

### Azure AD Configuration

Update `src/authConfig.js` with your Azure AD application details:

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

### SharePoint List Setup

Create a SharePoint list named "Tickets" with appropriate columns matching your field mapping requirements. Common columns:
- Title (subject)
- Description (multi-line text)
- User (single line text or person field)
- TicketNumber (single line text)
- AssignedTo (single line text or person field)

## Usage

### Creating Tickets from Emails

1. Click "Login with Microsoft" and authenticate
2. Click "Fetch Emails" to load recent messages
3. Select emails by clicking checkboxes
4. Click "Load SharePoint Fields" to fetch available columns
5. Upload technician CSV (format: Name,Email)
6. Map app fields to SharePoint columns
7. Assign technicians to selected emails
8. Click "Create Tickets" button

### Technician CSV Format

```csv
Name,Email
John Doe,john.doe@example.com
Jane Smith,jane.smith@example.com
Mike Johnson,mike.johnson@example.com
```

## Development

### Available Scripts

- `npm start`: Run development server (localhost:3000)
- `npm test`: Run test suite
- `npm run build`: Create production build
- `npm run eject`: Eject from Create React App (one-way operation)

### Code Style

- ES6+ JavaScript
- Functional React components with hooks
- Inline styles for component-specific styling
- External CSS for global and reusable styles

## Security Considerations

- OAuth 2.0 authentication via Microsoft Identity Platform
- Access tokens stored in sessionStorage (cleared on browser close)
- No client secret required (public client application)
- Redirect-based auth flow (no popup blocking issues)
- Scoped API permissions (least privilege principle)

## Troubleshooting

### Common Issues

**Login fails with AADSTS error:**
- Check Azure AD application configuration
- Verify redirect URI matches your origin
- Ensure API permissions are granted

**SharePoint API calls fail:**
- Verify the "Tickets" list exists at the root site
- Check Sites.ReadWrite.All permission is granted
- Confirm user has write access to SharePoint

**Token acquisition errors:**
- Clear browser cache and sessionStorage
- Re-authenticate with account picker
- Check network tab for detailed error responses

## Contributing

This is a private project. For questions or support, contact the repository owner.

## License

ISC

## Project Links

- Repository: https://github.com/aaronjt12/outlook-sync
- Issues: https://github.com/aaronjt12/outlook-sync/issues

## Changelog

### Latest (2025)
- Fixed MSAL authentication implementation
- Migrated from CDN to npm package for MSAL
- Added proper token acquisition after redirect
- Improved error handling and user feedback
- Created comprehensive documentation

### MVP (Initial Release)
- Email fetching from Outlook
- SharePoint ticket creation
- Field mapping functionality
- Technician assignment feature
