# Outlook Mail CLI Developer Guide

## Project Overview

The Outlook Mail CLI is a command-line application that interacts with Microsoft Outlook via the Microsoft Graph API. It allows users to perform various mail and folder operations directly from the command line.

## Development History

This project started as a simple authentication test application that could only list mail folders. It has since been expanded to include a more comprehensive set of features for interacting with Outlook mail.

### Initial Version (1.0.0)
- Basic authentication with Microsoft Graph API
- Ability to list top-level mail folders
- Ability to list child folders of a specific folder

### Current Version (1.0.0 - expanded)
- Added comprehensive email operations
- Added comprehensive folder operations
- Improved error handling and TypeScript typing
- Enhanced documentation

## Project Structure

```
outlook-auth-test-app-1.0/
├── dist/                # Compiled JavaScript files
├── node_modules/        # Dependencies
├── src/                 # Source code
│   ├── authHelper.ts    # Authentication with Microsoft Graph
│   ├── index.ts         # CLI commands and interface
│   └── mailService.ts   # Service for interacting with Microsoft Graph
├── .env                 # Environment variables (not in repo)
├── .env.sample          # Sample environment variables
├── package.json         # Project metadata and dependencies
├── tsconfig.json        # TypeScript configuration
└── README.md            # User documentation
```

## Key Components

### Authentication (authHelper.ts)

The application uses client credentials flow with Azure AD to authenticate with Microsoft Graph. This requires:
- A registered application in Azure AD
- Application permissions for Mail.Read or Mail.ReadWrite
- A client secret for authentication

See the AZURE_SETUP.md file for detailed setup instructions.

### Mail Service (mailService.ts)

This is the core of the application, containing:
- Interfaces for email, folder, and related data structures
- Methods for interacting with Microsoft Graph API endpoints
- Error handling for API communication

### Command Line Interface (index.ts)

This file defines all CLI commands using the Commander library, including:
- Command-line arguments and options
- Helper functions for displaying results
- Error handling for command execution

## Current Feature Set

### Email Operations
- List emails in a folder (`list-emails`)
- Read detailed email content (`read-email`)
- Move emails between folders (`move-email`)
- Copy emails between folders (`copy-email`)
- Create draft emails (`create-draft`)

### Folder Operations
- List top-level folders (`list-folders`)
- List child folders (`list-child-folders`)
- Create new folders (`create-folder`)
- Rename folders (`rename-folder`)
- Move folders between parents (`move-folder`)
- Copy folders (`copy-folder`) - with API support check

## Current Issues and Limitations

1. **API Limitations**: 
   - Folder copying may not be supported by Microsoft Graph API (error handling is in place)
   - No direct "archive" function in the API (archiving is typically implemented by moving emails to an archive folder)

2. **Missing Features**:
   - Email flagging/categorization
   - Email sending (currently only supports drafts)
   - Attachment handling
   - Reply/forward operations
   - Search functionality

3. **Edge Cases**:
   - Large attachment handling is not optimized
   - HTML email bodies are shown with simple tag stripping, which can be improved
   - Limited pagination for large mailboxes

## Planned Next Steps

### Short-term Improvements
1. Add support for email flagging and categorization
2. Add support for sending emails (not just drafts)
3. Improve HTML email rendering in the terminal
4. Add attachment download/upload capabilities

### Medium-term Improvements
1. Add search functionality for emails
2. Implement email reply and forward operations
3. Add support for meeting invitations
4. Implement better pagination for large mailboxes

### Long-term Vision
1. Create a TUI (Text User Interface) for easier mail navigation
2. Add email templates functionality
3. Support for rules and automated operations
4. Offline mode with synchronization

## Development Environment Setup

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd outlook-auth-test-app-1.0
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file based on `.env.sample` with your Azure credentials:
   ```
   TENANT_ID=your-tenant-id
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   ```

4. Build the project:
   ```bash
   npm run build
   ```

5. Test the application:
   ```bash
   node dist/index.js test-auth --user user@example.com
   ```

## Development Workflow

1. **Making Changes**:
   - Update TypeScript files in the `src/` directory
   - Run `npm run build` to compile
   - Test changes with appropriate commands

2. **Adding New Commands**:
   - Update `mailService.ts` with new methods for API interaction
   - Add new command definitions in `index.ts` using Commander's API
   - Update README.md with usage examples

3. **Testing**:
   - Always test commands with `--user` parameter (required for app-only authentication)
   - Build (with `npm run build`) before testing to ensure TypeScript compilation

## API Reference

For detailed information about the Microsoft Graph API endpoints:
- Mail API: https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview
- Message resource: https://learn.microsoft.com/en-us/graph/api/resources/message
- MailFolder resource: https://learn.microsoft.com/en-us/graph/api/resources/mailfolder

## Troubleshooting Common Issues

1. **Authentication Errors**:
   - Check `.env` file for correct credentials
   - Verify that app registration in Azure has correct permissions
   - Ensure admin consent has been granted for application permissions

2. **API Permission Errors**:
   - Check that the app has Mail.Read or Mail.ReadWrite permissions
   - For operations like moving emails, Mail.ReadWrite is required

3. **TypeScript Compilation Errors**:
   - Use proper type definitions, especially for error handling
   - Add type annotations for callback parameters, especially in array functions

## Contributing Guidelines

1. Make sure any new features follow the existing patterns:
   - Clear separation between API service and CLI interface
   - Proper TypeScript typing for all methods and parameters
   - Comprehensive error handling
   - Documentation for all new features

2. Update documentation when adding new features:
   - Add new commands to README.md
   - Document any API limitations or edge cases
   - Include usage examples

## Contact

For questions or issues about development, contact the project maintainer.
