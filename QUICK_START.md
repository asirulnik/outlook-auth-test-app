# Quick Start Guide

This guide provides a quick overview of how to get started with the Outlook Auth Test App.

## Prerequisites

- Node.js (v14 or higher)
- npm
- Azure account with admin access to register an application

## Setup in 5 Steps

1. **Clone the repository and install dependencies**

   ```bash
   cd /path/to/outlook-auth-test-app-1.0
   npm install
   ```

2. **Register an application in Azure**

   - Go to Azure Portal > Azure Active Directory > App registrations
   - Create a new registration
   - Grant Mail.Read and Mail.ReadBasic.All application permissions
   - Grant admin consent
   - Create a client secret
   
   For detailed instructions, see [AZURE_SETUP.md](./AZURE_SETUP.md)

3. **Configure your environment**

   Create a `.env` file in the root directory:

   ```
   TENANT_ID=your-tenant-id
   CLIENT_ID=your-application-client-id
   CLIENT_SECRET=your-client-secret
   ```

4. **Build the application**

   ```bash
   npm run build
   ```

5. **Run your first command**

   ```bash
   node dist/index.js test-auth --user user@yourdomain.com
   ```

   If successful, you should see: `Authentication successful! You are connected to Microsoft Graph API.`

## Common Commands

### Test Authentication

```bash
node dist/index.js test-auth --user user@yourdomain.com
```

### List Mail Folders

```bash
node dist/index.js list-folders --user user@yourdomain.com
```

### List Child Folders

```bash
node dist/index.js list-child-folders <folder-id> --user user@yourdomain.com
```

## Need Help?

- For Azure setup help, see [AZURE_SETUP.md](./AZURE_SETUP.md)
- For code explanation, see [CODE_EXPLANATION.md](./CODE_EXPLANATION.md)
- For detailed documentation, see [README.md](./README.md)
