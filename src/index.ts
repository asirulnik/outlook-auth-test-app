#!/usr/bin/env node
import { Command } from 'commander';
import { MailService, MailFolder } from './mailService';

// Create a new instance of the Command class
const program = new Command();

// Set up the program metadata
program
  .name('outlook-mail-cli')
  .description('CLI to authenticate with Microsoft Outlook and list mail folders')
  .version('1.0.0');

// Helper function to print folders in a tree structure
function printFolders(folders: MailFolder[], level = 0, prefix = '') {
  for (let i = 0; i < folders.length; i++) {
    const folder = folders[i];
    const isLast = i === folders.length - 1;
    const folderPrefix = isLast ? '└── ' : '├── ';
    const childPrefix = isLast ? '    ' : '│   ';
    
    // Print folder details
    console.log(`${prefix}${folderPrefix}${folder.displayName} ${formatFolderInfo(folder)}`);
    
    // If this folder has child folders, indicate with a message
    if (folder.childFolderCount > 0) {
      console.log(`${prefix}${childPrefix}<Folder has ${folder.childFolderCount} child folders. Use 'outlook-mail-cli list-child-folders ${folder.id}' to view.>`);
    }
  }
}

// Format folder information
function formatFolderInfo(folder: MailFolder): string {
  let info = `(ID: ${folder.id}`;
  
  if (folder.unreadItemCount !== undefined) {
    info += `, Unread: ${folder.unreadItemCount}`;
  }
  
  if (folder.totalItemCount !== undefined) {
    info += `, Total: ${folder.totalItemCount}`;
  }
  
  info += ')';
  return info;
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
      
      console.log(`\nMail Folders for ${options.user}:`);
      printFolders(folders);
      console.log('\n');
    } catch (error) {
      console.error('Error listing folders:', error);
      process.exit(1);
    }
  });

// Command to list child folders
program
  .command('list-child-folders <folderId>')
  .description('List child folders of a specific mail folder')
  .requiredOption('-u, --user <email>', 'Email address of the user (required for app-only authentication)')
  .action(async (folderId, options) => {
    try {
      const mailService = new MailService();
      const folders = await mailService.getChildFolders(folderId, options.user);
      
      console.log(`\nChild Folders for Folder ID: ${folderId} (User: ${options.user})`);
      printFolders(folders);
      console.log('\n');
    } catch (error) {
      console.error('Error listing child folders:', error);
      process.exit(1);
    }
  });

// Parse the command line arguments
program.parse();

// If no arguments provided, display help
if (process.argv.length < 3) {
  program.help();
}
