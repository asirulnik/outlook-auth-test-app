#!/usr/bin/env node
import { Command } from 'commander';
import { MailService, MailFolder, EmailMessage, EmailDetails, NewEmailDraft, NewMailFolder, EmailSearchOptions } from './mailService';
import { htmlToText } from './htmlToText';

// Create a new instance of the Command class
import * as fs from 'fs';

const program = new Command();

// Set up the program metadata
program
  .name('outlook-mail-cli')
  .description('CLI to interact with Microsoft Outlook mail')
  .version('1.0.0');

// Helper function to print folders in a tree structure
async function printFolders(folders: MailFolder[], mailService: MailService, userEmail: string, level = 0, prefix = '') {
  for (let i = 0; i < folders.length; i++) {
    const folder = folders[i];
    const isLast = i === folders.length - 1;
    const folderPrefix = isLast ? '└── ' : '├── ';
    const childPrefix = isLast ? '    ' : '│   ';
    
    // Get folder path
    const folderPath = folder.fullPath || await mailService.getFolderPath(folder.id, userEmail);
    
    // Print folder details with path instead of ID
    console.log(`${prefix}${folderPrefix}${folder.displayName} ${formatFolderInfo(folder, folderPath)}`);
    
    // If this folder has child folders, indicate with a message
    if (folder.childFolderCount > 0) {
      console.log(`${prefix}${childPrefix}<Folder has ${folder.childFolderCount} child folders. Use 'npx ts-node src/index.ts list-child-folders "${folderPath}" --user ${userEmail}' to view.>`);
    }
  }
}

// Format folder information
function formatFolderInfo(folder: MailFolder, folderPath: string): string {
  let info = `(Path: ${folderPath}`;
  
  if (folder.unreadItemCount !== undefined) {
    info += `, Unread: ${folder.unreadItemCount}`;
  }
  
  if (folder.totalItemCount !== undefined) {
    info += `, Total: ${folder.totalItemCount}`;
  }
  
  info += ')';
  return info;
}

// Helper function to print email messages
function printEmails(emails: EmailMessage[]) {
  console.log(`\nFound ${emails.length} emails:\n`);
  
  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    const fromName = email.from?.emailAddress.name || email.from?.emailAddress.address || 'Unknown Sender';
    const receivedDate = email.receivedDateTime ? new Date(email.receivedDateTime).toLocaleString() : 'Unknown Date';
    const hasAttachment = email.hasAttachments ? ' [Has Attachments]' : '';
    
    console.log(`${i + 1}. ${readStatus}${email.subject} - From: ${fromName} - ${receivedDate}${hasAttachment}`);
    console.log(`   ID: ${email.id}`);
    if (email.bodyPreview) {
      console.log(`   Preview: ${email.bodyPreview.substring(0, 100)}${email.bodyPreview.length > 100 ? '...' : ''}`);
    }
    console.log('');
  }
}

// Helper function to print email details
function printEmailDetails(email: EmailDetails) {
  console.log('\n==================================================');
  console.log(`Subject: ${email.subject}`);
  console.log(`From: ${email.from?.emailAddress.name || ''} <${email.from?.emailAddress.address || 'Unknown'}>`)
  
  if (email.toRecipients && email.toRecipients.length > 0) {
    console.log('To: ' + email.toRecipients.map(r => 
      `${r.emailAddress.name || ''} <${r.emailAddress.address}>`).join(', '));
  }
  
  if (email.ccRecipients && email.ccRecipients.length > 0) {
    console.log('CC: ' + email.ccRecipients.map(r => 
      `${r.emailAddress.name || ''} <${r.emailAddress.address}>`).join(', '));
  }
  
  if (email.receivedDateTime) {
    console.log(`Date: ${new Date(email.receivedDateTime).toLocaleString()}`);
  }
  
  if (email.attachments && email.attachments.length > 0) {
    console.log('\nAttachments:');
    email.attachments.forEach((attachment, i) => {
      const sizeInKB = Math.round(attachment.size / 1024);
      console.log(`${i + 1}. ${attachment.name} (${attachment.contentType}, ${sizeInKB} KB) - ID: ${attachment.id}`);
    });
  }
  
  console.log('\n--------------------------------------------------');
  if (email.body) {
    if (email.body.contentType === 'html') {
      console.log('Note: This is an HTML email. Plain text conversion shown:');
      // Use our enhanced HTML to text converter with formatting preservation
      const textContent = htmlToText(email.body.content, {
        wordwrap: 100, // Adjust based on terminal width
        preserveNewlines: true,
        tables: true,
        preserveHrefLinks: true,
        headingStyle: 'linebreak'
      });
      console.log(textContent);
    } else {
      console.log(email.body.content);
    }
  } else {
    console.log(email.bodyPreview || 'No content');
  }
  console.log('==================================================\n');
}

// Command to test authentication and display a success message
program
  .command('test-auth')
  .description('Test authentication with Microsoft Graph')
  .requiredOption('-u, --user <email>', 'Email address of the user (required for app-only authentication)')
  .action(async (options) => {
    try {
      const mailService = new MailService();
      // Try to get the top-level folders to test authentication
      await mailService.getMailFolders(options.user);
      console.log('Authentication successful! You are connected to Microsoft Graph API.');
    } catch (error) {
      console.error('Authentication failed:', error);
      process.exit(1);
    }
  });

// Command to list mail folders
program
  .command('list-folders')
  .description('List all top-level mail folders')
  .requiredOption('-u, --user <email>', 'Email address of the user (required for app-only authentication)')
  .action(async (options) => {
    try {
      const mailService = new MailService();
      const folders = await mailService.getMailFolders(options.user);
      
      // Build folder path map for all folders
      await mailService.buildFolderPathMap(options.user);
      
      console.log(`\nMail Folders for ${options.user}:`);
      await printFolders(folders, mailService, options.user);
      console.log('\n');
    } catch (error) {
      console.error('Error listing folders:', error);
      process.exit(1);
    }
  });

// Command to list child folders
program
  .command('list-child-folders <folderIdOrPath>')
  .description('List child folders of a specific mail folder')
  .requiredOption('-u, --user <email>', 'Email address of the user (required for app-only authentication)')
  .action(async (folderIdOrPath, options) => {
    try {
      const mailService = new MailService();
      
      // Resolve path if needed
      let folderPath = folderIdOrPath;
      if (!folderIdOrPath.startsWith('/')) {
        folderPath = await mailService.getFolderPath(folderIdOrPath, options.user);
      }
      
      const folders = await mailService.getChildFolders(folderIdOrPath, options.user);
      
      console.log(`\nChild Folders for Folder: ${folderPath} (User: ${options.user})`);
      await printFolders(folders, mailService, options.user);
      console.log('\n');
    } catch (error) {
      console.error('Error listing child folders:', error);
      process.exit(1);
    }
  });

// Command to list emails in a folder
program
  .command('list-emails <folderIdOrPath>')
  .description('List emails in a specific mail folder')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .option('-l, --limit <number>', 'Number of emails to retrieve', '25')
  .option('--before <date>', 'Only show emails before this date (YYYY-MM-DD)')
  .option('--after <date>', 'Only show emails after this date (YYYY-MM-DD)')
  .option('--previous <value>', 'Show emails from previous period (e.g., 7)')
  .option('--unit <unit>', 'Time unit for --previous (days, weeks, months, years)', 'days')
  .option('--search <query>', 'Search for emails containing the specified text')
  .option('--fields <fields>', 'Comma-separated list of fields to search (subject,body,from,recipients,all)', 'all')
  .action(async (folderIdOrPath, options) => {
    try {
      const mailService = new MailService();
      
      // Resolve path if needed
      let folderPath = folderIdOrPath;
      if (!folderIdOrPath.startsWith('/')) {
        folderPath = await mailService.getFolderPath(folderIdOrPath, options.user);
      }
      
      // Process search and date filters
      const searchOptions: EmailSearchOptions = {};
      
      if (options.before) {
        searchOptions.beforeDate = new Date(options.before);
        // Set time to end of day
        searchOptions.beforeDate.setHours(23, 59, 59, 999);
      }
      
      if (options.after) {
        searchOptions.afterDate = new Date(options.after);
        // Set time to start of day
        searchOptions.afterDate.setHours(0, 0, 0, 0);
      }
      
      if (options.previous && !isNaN(parseInt(options.previous))) {
        const value = parseInt(options.previous);
        const unit = options.unit as 'days' | 'weeks' | 'months' | 'years';
        
        if (['days', 'weeks', 'months', 'years'].includes(unit)) {
          searchOptions.previousPeriod = { value, unit };
        } else {
          console.warn(`Warning: Invalid time unit '${unit}'. Using 'days' instead.`);
          searchOptions.previousPeriod = { value, unit: 'days' };
        }
      }
      
      // Process search options
      if (options.search) {
        searchOptions.searchQuery = options.search;
        
        // Process search fields
        if (options.fields) {
          const validFields = ['subject', 'body', 'from', 'recipients', 'all'];
          const requestedFields = options.fields.split(',').map((f: string) => f.trim().toLowerCase());
          
          // Filter to only valid field values
          searchOptions.searchFields = requestedFields.filter((f: string) => 
            validFields.includes(f)
          ) as ('subject' | 'body' | 'from' | 'recipients' | 'all')[];
          
          // If no valid fields specified, default to 'all'
          if (searchOptions.searchFields.length === 0) {
            searchOptions.searchFields = ['all'];
          }
        } else {
          // Default to all fields
          searchOptions.searchFields = ['all'];
        }
      }
      
      const emails = await mailService.listEmails(
        folderIdOrPath, 
        options.user, 
        parseInt(options.limit), 
        Object.keys(searchOptions).length > 0 ? searchOptions : undefined
      );
      
      // Prepare filter description for output
      let filterDesc = '';
      if (searchOptions.beforeDate) {
        filterDesc += ` before ${searchOptions.beforeDate.toLocaleDateString()}`;
      }
      if (searchOptions.afterDate) {
        filterDesc += `${filterDesc ? ' and' : ''} after ${searchOptions.afterDate.toLocaleDateString()}`;
      }
      if (searchOptions.previousPeriod) {
        filterDesc = ` from previous ${searchOptions.previousPeriod.value} ${searchOptions.previousPeriod.unit}`;
      }
      if (searchOptions.searchQuery) {
        const searchDesc = searchOptions.searchFields?.includes('all') ? 
          `all fields` : 
          searchOptions.searchFields?.join(', ');
        
        filterDesc += `${filterDesc ? ' ' : ''} matching "${searchOptions.searchQuery}" in ${searchDesc}`;
      }
      
      console.log(`\nEmails in Folder: ${folderPath}${filterDesc} (User: ${options.user})`);
      printEmails(emails);

      // Print summary of results
      console.log(`\nFound ${emails.length} email(s) matching your criteria.`);
    } catch (error) {
      console.error('Error listing emails:', error);
      process.exit(1);
    }
  });

// Command to read a specific email
program
  .command('read-email <emailId>')
  .description('Read a specific email with all details')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .action(async (emailId, options) => {
    try {
      const mailService = new MailService();
      const email = await mailService.getEmail(emailId, options.user);
      
      printEmailDetails(email);
    } catch (error) {
      console.error('Error reading email:', error);
      process.exit(1);
    }
  });

// Command to move an email to another folder
program
  .command('move-email <emailId> <destinationFolderIdOrPath>')
  .description('Move an email to another folder')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .action(async (emailId, destinationFolderIdOrPath, options) => {
    try {
      const mailService = new MailService();
      
      // Resolve path if needed
      let folderPath = destinationFolderIdOrPath;
      if (!destinationFolderIdOrPath.startsWith('/')) {
        folderPath = await mailService.getFolderPath(destinationFolderIdOrPath, options.user);
      }
      
      await mailService.moveEmail(emailId, destinationFolderIdOrPath, options.user);
      
      console.log(`Email ${emailId} successfully moved to folder ${folderPath}`);
    } catch (error) {
      console.error('Error moving email:', error);
      process.exit(1);
    }
  });

// Command to copy an email to another folder
program
  .command('copy-email <emailId> <destinationFolderIdOrPath>')
  .description('Copy an email to another folder')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .action(async (emailId, destinationFolderIdOrPath, options) => {
    try {
      const mailService = new MailService();
      
      // Resolve path if needed
      let folderPath = destinationFolderIdOrPath;
      if (!destinationFolderIdOrPath.startsWith('/')) {
        folderPath = await mailService.getFolderPath(destinationFolderIdOrPath, options.user);
      }
      
      await mailService.copyEmail(emailId, destinationFolderIdOrPath, options.user);
      
      console.log(`Email ${emailId} successfully copied to folder ${folderPath}`);
    } catch (error) {
      console.error('Error copying email:', error);
      process.exit(1);
    }
  });

// Command to create a new draft email
program
  .command('create-draft')
  .description('Create a new draft email')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .requiredOption('-s, --subject <text>', 'Email subject')
  .requiredOption('-t, --to <emails>', 'Recipient email addresses (comma-separated)')
  .option('-c, --cc <emails>', 'CC recipient email addresses (comma-separated)')
  .option('-b, --bcc <emails>', 'BCC recipient email addresses (comma-separated)')
  .option('--html', 'Use HTML format for the body (default is plain text)')
  .option('-f, --file <path>', 'Path to a file containing the email body')
  .option('-m, --message <text>', 'Email body text (use this or --file)')
  .action(async (options) => {
    try {
      let bodyContent = '';
      
      // Check if body is provided via file or directly
      if (options.file) {
        try {
          bodyContent = fs.readFileSync(options.file, 'utf8');
        } catch (fsError) {
          console.error(`Error reading file: ${options.file}`, fsError);
          process.exit(1);
        }
      } else if (options.message) {
        bodyContent = options.message;
      } else {
        console.error('Error: Email body must be provided using --message or --file');
        process.exit(1);
      }
      
      // Parse recipients
      const toList = options.to.split(',').map((email: string) => ({
        emailAddress: { address: email.trim() }
      }));
      
      // Create draft object
      const draft: NewEmailDraft = {
        subject: options.subject,
        body: {
          contentType: options.html ? 'HTML' : 'Text',
          content: bodyContent
        },
        toRecipients: toList
      };
      
      // Add CC if provided
      if (options.cc) {
        draft.ccRecipients = options.cc.split(',').map((email: string) => ({
          emailAddress: { address: email.trim() }
        }));
      }
      
      // Add BCC if provided
      if (options.bcc) {
        draft.bccRecipients = options.bcc.split(',').map((email: string) => ({
          emailAddress: { address: email.trim() }
        }));
      }
      
      const mailService = new MailService();
      const result = await mailService.createDraft(draft, options.user);
      
      console.log(`Draft email created successfully with ID: ${result.id}`);
    } catch (error) {
      console.error('Error creating draft email:', error);
      process.exit(1);
    }
  });

// Command to create a new mail folder
program
  .command('create-folder <name>')
  .description('Create a new mail folder')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .option('-p, --parent <parentFolderIdOrPath>', 'Optional parent folder ID or path')
  .option('--hidden', 'Create the folder as hidden')
  .action(async (name, options) => {
    try {
      const newFolder: NewMailFolder = {
        displayName: name,
        isHidden: options.hidden || false
      };
      
      const mailService = new MailService();
      const result = await mailService.createFolder(newFolder, options.user, options.parent);
      
      // Get the path for the newly created folder
      const folderPath = await mailService.getFolderPath(result.id, options.user);
      
      console.log(`Folder "${name}" created successfully with path: ${folderPath}`);
    } catch (error) {
      console.error('Error creating folder:', error);
      process.exit(1);
    }
  });

// Command to rename a mail folder
program
  .command('rename-folder <folderIdOrPath> <newName>')
  .description('Rename a mail folder')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .action(async (folderIdOrPath, newName, options) => {
    try {
      const updatedFolder: Partial<NewMailFolder> = {
        displayName: newName
      };
      
      const mailService = new MailService();
      
      // Resolve path if needed and get current path for display
      let folderPath = folderIdOrPath;
      if (!folderIdOrPath.startsWith('/')) {
        folderPath = await mailService.getFolderPath(folderIdOrPath, options.user);
      }
      
      await mailService.updateFolder(folderIdOrPath, updatedFolder, options.user);
      
      // Get parent path
      const lastSlashIndex = folderPath.lastIndexOf('/');
      const parentPath = lastSlashIndex > 0 ? folderPath.substring(0, lastSlashIndex) : '';
      const newPath = parentPath + '/' + newName;
      
      console.log(`Folder ${folderPath} renamed to "${newName}" successfully`);
      console.log(`New path: ${newPath}`);
    } catch (error) {
      console.error('Error renaming folder:', error);
      process.exit(1);
    }
  });

// Command to move a folder to another parent folder
program
  .command('move-folder <folderIdOrPath> <destinationParentFolderIdOrPath>')
  .description('Move a folder to another parent folder')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .action(async (folderIdOrPath, destinationParentFolderIdOrPath, options) => {
    try {
      const mailService = new MailService();
      
      // Resolve paths if needed
      let sourceFolderPath = folderIdOrPath;
      if (!folderIdOrPath.startsWith('/')) {
        sourceFolderPath = await mailService.getFolderPath(folderIdOrPath, options.user);
      }
      
      let destinationFolderPath = destinationParentFolderIdOrPath;
      if (!destinationParentFolderIdOrPath.startsWith('/')) {
        destinationFolderPath = await mailService.getFolderPath(destinationParentFolderIdOrPath, options.user);
      }
      
      await mailService.moveFolder(folderIdOrPath, destinationParentFolderIdOrPath, options.user);
      
      // Get folder name from source path
      const folderName = sourceFolderPath.substring(sourceFolderPath.lastIndexOf('/') + 1);
      const newPath = destinationFolderPath + '/' + folderName;
      
      console.log(`Folder ${sourceFolderPath} successfully moved to ${destinationFolderPath}`);
      console.log(`New path: ${newPath}`);
    } catch (error) {
      console.error('Error moving folder:', error);
      process.exit(1);
    }
  });

// Command to copy a folder to another parent folder
program
  .command('copy-folder <folderIdOrPath> <destinationParentFolderIdOrPath>')
  .description('Copy a folder to another parent folder (may not be supported by the API)')
  .requiredOption('-u, --user <email>', 'Email address of the user')
  .action(async (folderIdOrPath, destinationParentFolderIdOrPath, options) => {
    try {
      const mailService = new MailService();
      
      // Resolve paths if needed
      let sourceFolderPath = folderIdOrPath;
      if (!folderIdOrPath.startsWith('/')) {
        sourceFolderPath = await mailService.getFolderPath(folderIdOrPath, options.user);
      }
      
      let destinationFolderPath = destinationParentFolderIdOrPath;
      if (!destinationParentFolderIdOrPath.startsWith('/')) {
        destinationFolderPath = await mailService.getFolderPath(destinationParentFolderIdOrPath, options.user);
      }
      
      await mailService.copyFolder(folderIdOrPath, destinationParentFolderIdOrPath, options.user);
      
      // Get folder name from source path
      const folderName = sourceFolderPath.substring(sourceFolderPath.lastIndexOf('/') + 1);
      const newPath = destinationFolderPath + '/' + folderName;
      
      console.log(`Folder ${sourceFolderPath} successfully copied to ${destinationFolderPath}`);
      console.log(`Copy path: ${newPath}`);
    } catch (error: unknown) {
      const err = error as { message?: string };
      if (err.message?.includes('not supported')) {
        console.error('Error: Folder copying is not supported by the Microsoft Graph API');
      } else {
        console.error('Error copying folder:', error);
      }
      process.exit(1);
    }
  });

// Parse the command line arguments
program.parse();

// If no arguments provided, display help
if (process.argv.length < 3) {
  program.help();
}
