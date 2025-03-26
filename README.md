# Outlook Auth Test App

A simple CLI application that authenticates to Microsoft Outlook via the Microsoft Graph API and lists mail folders.

## Features

- Authenticate to Microsoft Graph API using client credentials flow
- List all top-level mail folders
- List child folders of a specific mail folder
- Support for accessing mailboxes by specifying a user email (with proper permissions)

## Prerequisites

- Node.js (v14 or higher)
- npm or yarn
- Microsoft Azure account with proper permissions
- Registered application in Azure AD

## Setup

1. Clone this repository or download the source code
2. Install dependencies:
   ```
   npm install
   ```
3. Create a `.env` file in the root directory with the following variables:
   ```
   TENANT_ID=your-tenant-id
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   ```

## Azure AD Application Setup

1. Register a new application in the Azure portal

2. Set the proper API permissions for application (client credentials) flow:
   - Microsoft Graph > Application permissions > Mail.Read (or Mail.ReadWrite for more functionality)
   - Microsoft Graph > Application permissions > Mail.ReadBasic.All (to read mail from all users)

3. **Important:** Grant admin consent for these permissions by clicking the "Grant admin consent for [your organization]" button

4. Create a client secret:
   - Navigate to **Certificates & secrets**
   - Click **New client secret**
   - Set a description and expiration
   - Copy the secret value immediately (you won't be able to see it again)

5. Update your .env file with your tenant ID, client ID, and client secret

## Usage

First, build the application:

```
npm run build
```

### Test Authentication

```
npx ts-node src/index.ts test-auth --user user@example.com
```
or
```
npm run build
node dist/index.js test-auth --user user@example.com
```

### List Top-Level Mail Folders

```
npx ts-node src/index.ts list-folders --user user@example.com
```
or
```
node dist/index.js list-folders --user user@example.com
```

### List Child Folders

```
npx ts-node src/index.ts list-child-folders <folder-id> --user user@example.com
```
or
```
node dist/index.js list-child-folders <folder-id> --user user@example.com
```

## Common Issues

- **Authentication failed**: Make sure your client ID, tenant ID, and client secret are correct in the .env file
- **Permission denied**: Ensure you've granted the proper application permissions to your application and provided admin consent in the Azure portal
- **401 Unauthorized**: Check that your client secret hasn't expired and that your application has the proper permissions
- **"...is only valid with delegated authentication flow"**: This app uses application permissions (client credentials flow) and requires a specific user email for all operations. Make sure to use the `--user` parameter
- **"User does not exist or one of its dependencies"**: Check that the email address you're using exists and is accessible to your application

## License

ISC
