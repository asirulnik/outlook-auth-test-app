# Azure AD Setup for Outlook Auth Test App

This document provides detailed steps for setting up the Azure AD application registration required for this application to work with client credentials flow (app-only authentication).

## Prerequisites

- Admin access to an Azure AD tenant
- At least one user mailbox that you want to access through the application

## Step 1: Register a New Application

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **+ New registration**
4. Fill in the form:
   - **Name**: Outlook Mail CLI App
   - **Supported account types**: Select "Accounts in this organizational directory only" (single tenant)
   - **Redirect URI**: Leave blank (not required for app-only authentication)
5. Click **Register**

## Step 2: Note Application Details

After registration, you'll see the app's overview page. Note the following information:

- **Application (client) ID**: This will be used as the `CLIENT_ID` in your .env file
- **Directory (tenant) ID**: This will be used as the `TENANT_ID` in your .env file

## Step 3: Add API Permissions

1. From your app's overview page, select **API permissions** from the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions**
5. Find and select the following permissions:
   - **Mail.Read** (allows reading mail)
   - **Mail.ReadBasic.All** (allows reading basic mail properties of all users)
   - **MailboxSettings.Read** (optional, for reading mailbox settings)

   You may need different permissions depending on your specific requirements.

6. Click **Add permissions**

## Step 4: Grant Admin Consent

1. After adding permissions, you'll see them listed with a status of "Not granted"
2. Click the **Grant admin consent for [Your Organization]** button
3. Click **Yes** to confirm

This step is critical - without admin consent, your application cannot use the permissions.

## Step 5: Create a Client Secret

1. From your app's overview page, select **Certificates & secrets** from the left menu
2. Under **Client secrets**, click **+ New client secret**
3. Add a description and select an expiration period
4. Click **Add**
5. **IMPORTANT**: Copy the secret value immediately and store it securely. 
   This will be used as the `CLIENT_SECRET` in your .env file.
   You won't be able to see this value again after leaving this page.

## Step 6: Update the Application's .env File

Create or update the .env file in your application directory with the following values:

```
TENANT_ID=your-tenant-id
CLIENT_ID=your-application-client-id
CLIENT_SECRET=your-client-secret
```

## Testing the Setup

To test if your setup works correctly:

1. Build the application:
   ```
   npm run build
   ```

2. Test authentication for a specific user:
   ```
   node dist/index.js test-auth --user andrew@sirulnik-law.com
   ```

3. If successful, try listing the mail folders:
   ```
   
   ```

## Troubleshooting

If you encounter errors:

- **401 Unauthorized**: Check your client ID, tenant ID, and client secret
- **403 Forbidden**: Verify that admin consent has been granted for the permissions
- **Not found**: Ensure the user email exists and is accessible to your application
- **Bad Request**: Make sure you're providing a user email with the `--user` parameter, as this application uses app-only authentication which requires specifying a user
