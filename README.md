# Outlook Mail API Client

A lightweight TypeScript client for interacting with Microsoft Outlook Mail API. This tool allows you to access and manage your Outlook mailbox through simple commands.

## Installation

1. Clone this repository
2. Install dependencies:
   ```
   npm install
   ```
3. Configure authentication (see Authentication section below)

## Usage

This tool uses a single-command approach where you specify the operation and parameters in one command line instruction.

### Basic syntax:

```
npx ts-node src/index.ts <command> [options]
```

### Available commands:

- **list-folders**: Display all mail folders
  ```
  npx ts-node src/index.ts list-folders --user your-email@example.com
  ```

- **list-child-folders**: View subfolders within a specific folder
  ```
  npx ts-node src/index.ts list-child-folders "/Inbox" --user your-email@example.com
  ```

- **list-emails**: List messages in a folder with search and date filtering options
  ```
  npx ts-node src/index.ts list-emails "/Inbox" --user your-email@example.com [--limit 10] [--search "keyword"] [--fields subject,body] [--before 2023-12-31] [--after 2023-01-01] [--previous 7 --unit days]
  ```

- **read-email**: Retrieve a specific email message
  ```
  npx ts-node src/index.ts read-email <message-id> --user your-email@example.com
  ```

- **move-email**: Move an email to another folder
  ```
  npx ts-node src/index.ts move-email <message-id> "/Archive" --user your-email@example.com
  ```

- **copy-email**: Copy an email to another folder
  ```
  npx ts-node src/index.ts copy-email <message-id> "/Important Emails" --user your-email@example.com
  ```

- **create-draft**: Create a new email
  ```
  npx ts-node src/index.ts create-draft --to recipient@example.com --subject "Hello" --message "Email content" --user your-email@example.com
  ```

- **create-folder**: Create a new mail folder
  ```
  npx ts-node src/index.ts create-folder "New Folder" --user your-email@example.com [--parent "/Clients"]
  ```

- **rename-folder**: Rename an existing folder
  ```
  npx ts-node src/index.ts rename-folder "/Clients/Old Name" "New Name" --user your-email@example.com
  ```

- **move-folder**: Move a folder to another parent folder
  ```
  npx ts-node src/index.ts move-folder "/Folder" "/New Parent" --user your-email@example.com
  ```

## Authentication

This tool requires OAuth authentication with Microsoft Graph API. To configure:

1. Register an application in Azure Portal (https://portal.azure.com)
2. Add Microsoft Graph API permissions for Mail.Read, Mail.Send, etc.
3. Create a `.env` file with the following:
   ```
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   TENANT_ID=your_tenant_id
   REDIRECT_URI=http://localhost:3000/auth/callback
   ```

The first time you run a command, you'll need to authenticate through the browser.

## Folder Path Support

This tool supports both folder IDs and plain-text folder paths:

- Use natural folder paths like "/Inbox" or "/Clients/ProjectX" 
- Paths are case-insensitive
- Paths must start with a forward slash (/)
- Nested paths like "/ParentFolder/ChildFolder" work seamlessly
- The tool will automatically convert between paths and IDs as needed

## Search and Filter Options

When listing emails, you can use powerful search and date filtering capabilities:

### Full-Text Search Options

- **--search "QUERY"**: Search for emails containing the specified text
- **--fields FIELDS**: Comma-separated list of fields to search in (default: all)
  - Available fields: subject, body, from, recipients, all
  - Example: --fields subject,body

### Date Filtering Options

- **--before YYYY-MM-DD**: Show only emails received before the specified date
- **--after YYYY-MM-DD**: Show only emails received after the specified date
- **--previous VALUE**: Show emails from the previous period (e.g., 7)
- **--unit UNIT**: Time unit for --previous (days, weeks, months, years)

These options can be used individually or combined for powerful, targeted queries:

```
# Search for emails containing "project update" in the subject or body from the last 30 days
npx ts-node src/index.ts list-emails "/Inbox" --user your-email@example.com --search "project update" --fields subject,body --previous 30 --unit days

# Search for emails from a specific sender in Q1 2023
npx ts-node src/index.ts list-emails "/Archive" --user your-email@example.com --search "john.doe@example.com" --fields from --after 2023-01-01 --before 2023-03-31

# Search for all mentions of "contract" in any field during the previous year
npx ts-node src/index.ts list-emails "/Clients" --user your-email@example.com --search "contract" --previous 1 --unit years
```

The search is performed server-side using Microsoft Graph API's search capabilities, which provides fast and efficient results even for large mailboxes.

## Examples

List all folders in your mailbox:
```
npx ts-node src/index.ts list-folders --user your-email@example.com
```

Search for emails containing "invoice" in the subject from the last week:
```
npx ts-node src/index.ts list-emails "/Inbox" --user your-email@example.com --search "invoice" --fields subject --previous 7 --unit days
```

Move an email to a specific folder:
```
npx ts-node src/index.ts move-email AAMkAGFmOThiYzZj-GFmOThiYzZj-wEKAAAA "/Archive/2023" --user your-email@example.com
```

Create a new folder inside another folder:
```
npx ts-node src/index.ts create-folder "Project XYZ" --parent "/Clients" --user your-email@example.com
```

## License

MIT
