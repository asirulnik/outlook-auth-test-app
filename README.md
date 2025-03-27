# Outlook Auth Test App

A CLI application that interacts with Microsoft Outlook via the Microsoft Graph API to manage emails and folders.

## Features

- Authenticate to Microsoft Graph API using client credentials flow
- Mail Folder Operations:
  - List all top-level mail folders
  - List child folders of a specific mail folder
  - Create new folders
  - Rename folders
  - Move folders to different parent folders
  - Copy folders (if supported by the API)
- Email Operations:
  - List emails in a folder
  - Read email content
  - Move emails between folders
  - Copy emails between folders
  - Create new email drafts
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
   - Microsoft Graph > Application permissions > Mail.Send (to send emails)

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

### Mail Folder Operations

#### List Top-Level Mail Folders

```
npx ts-node src/index.ts list-folders --user user@example.com
```

#### List Child Folders

```
npx ts-node src/index.ts list-child-folders <folder-id> --user user@example.com
```

#### Create a New Folder

```
npx ts-node src/index.ts create-folder "Folder Name" --user user@example.com
```

With a parent folder:
```
npx ts-node src/index.ts create-folder "Folder Name" --user user@example.com --parent <parent-folder-id>
```

#### Rename a Folder

```
npx ts-node src/index.ts rename-folder <folder-id> "New Folder Name" --user user@example.com
```

#### Move a Folder

```
npx ts-node src/index.ts move-folder <folder-id> <destination-parent-folder-id> --user user@example.com
```

#### Copy a Folder (if supported by the API)

```
npx ts-node src/index.ts copy-folder <folder-id> <destination-parent-folder-id> --user user@example.com
```

### Email Operations

#### List Emails in a Folder

```
npx ts-node src/index.ts list-emails <folder-id> --user user@example.com
```

With a custom limit:
```
npx ts-node src/index.ts list-emails <folder-id> --user user@example.com --limit 50
```

#### Read an Email

```
npx ts-node src/index.ts read-email <email-id> --user user@example.com
```

#### Move an Email to Another Folder

```
npx ts-node src/index.ts move-email <email-id> <destination-folder-id> --user user@example.com
```

#### Copy an Email to Another Folder

```
npx ts-node src/index.ts copy-email <email-id> <destination-folder-id> --user user@example.com
```

#### Create a Draft Email

```
npx ts-node src/index.ts create-draft --user user@example.com --subject "Email Subject" --to "recipient@example.com" --message "Email body text"
```

With CC and BCC:
```
npx ts-node src/index.ts create-draft --user user@example.com --subject "Email Subject" --to "recipient@example.com" --cc "cc@example.com" --bcc "bcc@example.com" --message "Email body text"
```

Using HTML format:
```
npx ts-node src/index.ts create-draft --user user@example.com --subject "Email Subject" --to "recipient@example.com" --html --message "<h1>Hello</h1><p>This is HTML email</p>"
```

Using a file for the body:
```
npx ts-node src/index.ts create-draft --user user@example.com --subject "Email Subject" --to "recipient@example.com" --file path/to/email-body.txt
```

## Common Issues

- **Authentication failed**: Make sure your client ID, tenant ID, and client secret are correct in the .env file
- **Permission denied**: Ensure you've granted the proper application permissions to your application and provided admin consent in the Azure portal
- **401 Unauthorized**: Check that your client secret hasn't expired and that your application has the proper permissions
- **"...is only valid with delegated authentication flow"**: This app uses application permissions (client credentials flow) and requires a specific user email for all operations. Make sure to use the `--user` parameter
- **"User does not exist or one of its dependencies"**: Check that the email address you're using exists and is accessible to your application
- **"Folder copying is not supported by the Microsoft Graph API"**: The Microsoft Graph API may not support copying folders. In that case, you'll need to implement a custom solution if you need this functionality.

## HTML to Plain Text Converter

This application includes an enhanced HTML to plain text converter that preserves formatting elements when displaying HTML emails. The converter supports:

- Maintaining paragraph structure with proper spacing
- Converting HTML lists (both ordered and unordered) to plain text lists
- Preserving table structure
- Converting links with displayed URLs
- Handling blockquotes, pre-formatted text, and horizontal rules
- Formatting text styles (bold, italic, underline, etc.)
- Configurable word wrapping
- Maintaining heading hierarchy

### Testing the HTML to Plain Text Converter

You can test the HTML to text converter directly with HTML files:

```
npm run test-html2text test-email-1.html
```

or

```
npx ts-node src/test-converter.ts path/to/html/file.html
```

This will convert the HTML file to plain text and save the output as a .txt file with the same name.

### Customizing the Converter

The converter accepts the following options:

- `wordwrap`: Character limit before wrapping (number or false to disable)
- `preserveNewlines`: Whether to keep existing newlines (boolean)
- `tables`: Whether to format tables (boolean)
- `uppercaseHeadings`: Whether to convert headings to uppercase (boolean)
- `preserveHrefLinks`: Whether to include URLs in brackets after link text (boolean)
- `bulletIndent`: Indentation for bullets (number)
- `listIndent`: Indentation for lists (number)
- `headingStyle`: How to format headings ('underline', 'linebreak', or 'hashify')
- `maxLineLength`: Maximum line length (number)

## License

ISC
