# Code Explanation

This document explains the structure and code of the Outlook Auth Test App.

## Overview

This application is a command-line interface (CLI) tool that:

1. Authenticates to Microsoft Graph API using client credentials flow (app-only authentication)
2. Lists mail folders for a specified user's mailbox
3. Lists child folders for a specified parent folder
4. Lists emails in folders with advanced filtering options
5. Reads emails with options to filter out quoted content
6. Moves and copies emails between folders
7. Creates, renames, and moves folders

The app uses TypeScript and leverages the following key libraries:

- `@azure/identity`: Provides the authentication capability for Microsoft Graph
- `@microsoft/microsoft-graph-client`: Client library for interacting with Microsoft Graph API
- `commander`: For building the CLI interface
- `dotenv`: For loading environment variables from a .env file
- `isomorphic-fetch`: For providing the fetch API in Node.js environments

## File Structure

- `src/authHelper.ts`: Handles the authentication logic
- `src/mailService.ts`: Provides methods for interacting with mail folders and emails
- `src/htmlToText.ts`: Utilities for converting HTML email bodies to text and filtering quoted content
- `src/index.ts`: Main CLI application logic
- `src/mcp-server.ts`: MCP server implementation for integration with Claude and other LLMs
- `.env`: Contains authentication credentials

## Authentication Flow

The application uses the client credentials flow (app-only authentication) with a client secret. This is different from delegated authentication in that:

1. It authenticates as the application itself, not on behalf of a user
2. It requires pre-configured application permissions in Azure AD
3. It requires admin consent for these permissions
4. It cannot use endpoints like `/me` that require user context

The authentication process works as follows:

1. The `ClientSecretCredential` class from `@azure/identity` is used to create a credential object with tenant ID, client ID, and client secret.
2. A `TokenCredentialAuthenticationProvider` is created using this credential.
3. The Microsoft Graph client is initialized with this authentication provider.
4. When making API calls, the credential automatically acquires access tokens.

## Key Components

### AuthHelper (authHelper.ts)

This module handles authentication to Microsoft Graph:
- Loads credentials from environment variables
- Creates the `ClientSecretCredential` with tenant ID, client ID, and client secret
- Creates an authentication provider that uses the `https://graph.microsoft.com/.default` scope
- Initializes and returns a Microsoft Graph client

### MailService (mailService.ts)

This service provides several key methods:
- `getMailFolders(userEmail)`: Gets top-level mail folders for a specified user
- `getChildFolders(folderIdOrWellKnownName, userEmail)`: Gets child folders for a specific parent folder
- `listEmails(folderIdOrWellKnownName, userEmail, limit, searchOptions)`: Lists emails in a folder with advanced filtering
- `getEmail(emailId, userEmail, hideQuotedContent)`: Gets detailed email content with option to filter quoted content
- `moveEmail(emailId, destinationFolderId, userEmail)`: Moves an email to another folder
- `copyEmail(emailId, destinationFolderId, userEmail)`: Copies an email to another folder
- `createDraft(draft, userEmail)`: Creates a new email draft
- `createFolder(newFolder, userEmail, parentFolderId)`: Creates a new mail folder
- `updateFolder(folderId, updatedFolder, userEmail)`: Updates a folder's properties
- `moveFolder(folderId, destinationParentFolderId, userEmail)`: Moves a folder to another parent folder

Each method builds the appropriate API endpoint for the Microsoft Graph API and makes the request.

The service defines several interfaces that represent the structure of mail folder objects, email messages, email details, and search options.

### HTML To Text Converter (htmlToText.ts)

This module provides functionality for converting HTML email content to plain text:
- Preserves basic formatting like paragraphs, lists, and tables
- Handles email signatures and quoted content
- Can identify and mark quoted prior emails with separator lines
- Supports filtering quoted content for cleaner email viewing

### CLI Application (index.ts)

This is the main entry point that:
- Sets up the CLI commands using Commander
- Implements functions for all mail operations (listing folders, listing emails, reading emails, etc.)
- Provides formatted output for the command results
- Supports advanced filtering and display options

The CLI requires a user email for all commands since the app uses app-only authentication, which doesn't have a default user context.

### MCP Server (mcp-server.ts)

This module implements a Model Context Protocol (MCP) server:
- Exposes mail operations as tools for LLMs like Claude
- Handles communication via the MCP protocol
- Provides tools for listing folders, listing emails, reading emails, etc.
- Supports the same advanced features as the CLI (including filtering quoted content)

## Permissions

The application requires the following Microsoft Graph permissions:
- `Mail.Read` (application permission): To read mail folders and email messages
- `Mail.ReadBasic.All` (application permission): To read basic mail properties of all users
- `Mail.ReadWrite` (application permission): To create, update, move, and delete mail folders and messages

These permissions must be explicitly granted by an administrator in the Azure AD portal.

## Important Implementation Notes

1. **User Email Requirement**:
   - All commands require a user email parameter (`--user`)
   - This is because app-only authentication doesn't have a user context, so you must specify which user's data to access

2. **Error Handling**:
   - The application has robust error handling that outputs errors to the console
   - Common errors include authentication failures, permissions issues, and invalid user emails

3. **Folder Format**:
   - Folders are displayed in a hierarchical format with information about child folders
   - For folders with children, a message is shown instructing how to list the child folders

4. **Folder Path Resolution**:
   - The application supports both folder IDs and human-readable paths
   - Paths like "/Inbox" or "/Archive/2023" are automatically resolved to folder IDs
   - A caching system improves performance by mapping paths to IDs

5. **Email Content Handling**:
   - HTML email content is converted to readable plain text
   - Quoted content can be identified and optionally filtered out
   - Original content is preserved when filtering for reference

6. **Search and Filtering**:
   - Powerful search options include full-text search across multiple fields
   - Date filtering supports specific dates and relative periods
   - Filtering can be combined for detailed queries

7. **API Endpoints**:
   - The application uses various Microsoft Graph API endpoints:
     - `/users/{userEmail}/mailFolders` for accessing mail folders
     - `/users/{userEmail}/mailFolders/{folderId}/childFolders` for nested folders
     - `/users/{userEmail}/mailFolders/{folderId}/messages` for listing emails
     - `/users/{userEmail}/messages/{messageId}` for detailed email access
     - Various other endpoints for moving, copying, and creating items

## Advanced Features

### Email Body Handling

The application includes sophisticated email body processing:

1. **HTML to Text Conversion**:
   - Preserves formatting including lists, tables, and paragraphs
   - Maintains indentation and structure
   - Handles hyperlinks and basic styling

2. **Quoted Content Detection**:
   - Identifies common email quoting patterns using multiple heuristics
   - Detects forwarded messages and reply chains
   - Uses HTML structure analysis and text pattern matching
   - Marks quoted content with separator lines

3. **Content Filtering**:
   - Can show only the primary content of an email, removing quoted portions
   - Preserves original content for reference
   - Adds indicators when content has been removed

### MCP Integration

The MCP server implementation enables AI assistants to:

1. Access and navigate mailbox folders
2. Search and filter emails with complex criteria
3. Read email content with or without quoted portions
4. Move and organize emails and folders

This provides AI assistants with the ability to help users manage their email more effectively while maintaining control over sensitive content.

## Future Enhancements

Potential future enhancements include:

1. Implementing email sending capability
2. Adding attachment handling (download, upload, view)
3. Adding calendar integration
4. Implementing contact management
5. Adding support for different authentication methods
6. Building a web interface for easier interaction
7. Implementing advanced AI summarization for email content

Each of these would require appropriate permissions and potentially additional endpoints from the Microsoft Graph API.
