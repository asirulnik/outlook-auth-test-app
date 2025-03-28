import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { MailService, EmailSearchOptions } from './mailService';
import { htmlToText } from './htmlToText';

/**
 * Outlook MCP Server Implementation
 * Provides tools for interacting with Microsoft Outlook mail
 */

async function main() {
  console.log('Starting Outlook MCP Server...');
  
  // Create the MCP server
  const server = new McpServer({
    name: 'Outlook MCP Server',
    version: '1.0.0'
  });

  // 1. List mail folders tool
  server.tool(
    'list-mail-folders',
    { 
      userEmail: z.string().email()
    },
    async ({ userEmail }) => {
      try {
        const mailService = new MailService();
        const folders = await mailService.getMailFolders(userEmail);
        
        return {
          content: [{ 
            type: 'text', 
            text: JSON.stringify(folders, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error listing mail folders: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // 2. List child folders tool
  server.tool(
    'list-child-folders',
    { 
      userEmail: z.string().email(),
      folderId: z.string()
    },
    async ({ userEmail, folderId }) => {
      try {
        const mailService = new MailService();
        const folders = await mailService.getChildFolders(folderId, userEmail);
        
        return {
          content: [{ 
            type: 'text', 
            text: JSON.stringify(folders, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error listing child folders: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // 3. List emails tool with options for full bodies and hiding quoted content
  server.tool(
    'list-emails',
    { 
      userEmail: z.string().email(),
      folderId: z.string(),
      limit: z.number().min(1).max(100).optional().default(25),
      includeBodies: z.boolean().optional().default(false),
      hideQuotedContent: z.boolean().optional().default(false),
      searchOptions: z.object({
        beforeDate: z.string().optional(),
        afterDate: z.string().optional(),
        previousPeriod: z.object({
          value: z.number(),
          unit: z.enum(['days', 'weeks', 'months', 'years'])
        }).optional(),
        searchQuery: z.string().optional(),
        searchFields: z.array(z.enum(['subject', 'body', 'from', 'recipients', 'all'])).optional()
      }).optional()
    },
    async ({ userEmail, folderId, limit, includeBodies, hideQuotedContent, searchOptions }) => {
      try {
        const mailService = new MailService();
        
        // Process search options
        const searchParams: EmailSearchOptions = {
          includeBodies,
          hideQuotedContent
        };
        
        if (searchOptions) {
          // Process date options
          if (searchOptions.beforeDate) {
            searchParams.beforeDate = new Date(searchOptions.beforeDate);
          }
          
          if (searchOptions.afterDate) {
            searchParams.afterDate = new Date(searchOptions.afterDate);
          }
          
          // Process previous period
          if (searchOptions.previousPeriod) {
            searchParams.previousPeriod = searchOptions.previousPeriod;
          }
          
          // Process search query and fields
          if (searchOptions.searchQuery) {
            searchParams.searchQuery = searchOptions.searchQuery;
            searchParams.searchFields = searchOptions.searchFields;
          }
        }
        
        const emails = await mailService.listEmails(folderId, userEmail, limit, searchParams);
        
        return {
          content: [{ 
            type: 'text', 
            text: JSON.stringify(emails, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error listing emails: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // 4. Read email tool with option to hide quoted content
  server.tool(
    'read-email',
    { 
      userEmail: z.string().email(),
      emailId: z.string(),
      convertHtmlToText: z.boolean().optional().default(true),
      hideQuotedContent: z.boolean().optional().default(false),
      htmlToTextOptions: z.object({
        wordwrap: z.union([z.number(), z.boolean()]).optional(),
        preserveNewlines: z.boolean().optional(),
        tables: z.boolean().optional(),
        preserveHrefLinks: z.boolean().optional(),
        headingStyle: z.enum(['underline', 'linebreak', 'hashify']).optional()
      }).optional()
    },
    async ({ userEmail, emailId, convertHtmlToText, hideQuotedContent, htmlToTextOptions }) => {
      try {
        const mailService = new MailService();
        const email = await mailService.getEmail(emailId, userEmail, hideQuotedContent);
        
        // Process HTML content if needed
        if (convertHtmlToText && 
            email.body && 
            email.body.contentType === 'html') {
          
          // Apply HTML to text conversion
          const defaultOptions = {
            wordwrap: 100,
            preserveNewlines: true,
            tables: true,
            preserveHrefLinks: true,
            headingStyle: 'linebreak' as const
          };
          
          const options = {
            ...defaultOptions,
            ...htmlToTextOptions
          };
          
          // Convert HTML to plain text
          email.body.plainTextContent = htmlToText(email.body.content, options);
          
          // If we have the original content (before removing quoted parts), convert that too
          if (hideQuotedContent && email.body.originalContent) {
            email.body.originalPlainTextContent = htmlToText(email.body.originalContent, options);
          }
        }
        
        return {
          content: [{ 
            type: 'text', 
            text: JSON.stringify(email, null, 2)
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error reading email: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // 5. Move email tool
  server.tool(
    'move-email',
    { 
      userEmail: z.string().email(),
      emailId: z.string(),
      destinationFolderId: z.string()
    },
    async ({ userEmail, emailId, destinationFolderId }) => {
      try {
        const mailService = new MailService();
        const result = await mailService.moveEmail(emailId, destinationFolderId, userEmail);
        
        return {
          content: [{ 
            type: 'text', 
            text: `Email ${emailId} successfully moved to folder ${destinationFolderId}`
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error moving email: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // 6. Create email draft tool
  server.tool(
    'create-draft',
    { 
      userEmail: z.string().email(),
      subject: z.string(),
      body: z.string(),
      isHtml: z.boolean().optional().default(false),
      to: z.array(z.string().email()),
      cc: z.array(z.string().email()).optional(),
      bcc: z.array(z.string().email()).optional()
    },
    async ({ userEmail, subject, body, isHtml, to, cc, bcc }) => {
      try {
        const mailService = new MailService();
        
        const draft = {
          subject,
          body: {
            contentType: isHtml ? 'HTML' : 'Text',
            content: body
          },
          toRecipients: to.map(email => ({
            emailAddress: { address: email }
          }))
        };
        
        // Add CC if provided
        if (cc && cc.length > 0) {
          draft.ccRecipients = cc.map(email => ({
            emailAddress: { address: email }
          }));
        }
        
        // Add BCC if provided
        if (bcc && bcc.length > 0) {
          draft.bccRecipients = bcc.map(email => ({
            emailAddress: { address: email }
          }));
        }
        
        const result = await mailService.createDraft(draft, userEmail);
        
        return {
          content: [{ 
            type: 'text', 
            text: `Draft email created successfully with ID: ${result.id}`
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error creating draft email: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // 7. Convert HTML to plain text tool
  server.tool(
    'convert-html-to-text',
    { 
      html: z.string(),
      options: z.object({
        wordwrap: z.union([z.number(), z.boolean()]).optional(),
        preserveNewlines: z.boolean().optional(),
        tables: z.boolean().optional(),
        preserveHrefLinks: z.boolean().optional(),
        headingStyle: z.enum(['underline', 'linebreak', 'hashify']).optional(),
        bulletIndent: z.number().optional(),
        listIndent: z.number().optional(),
        maxLineLength: z.number().optional(),
        hideQuotedContent: z.boolean().optional()
      }).optional()
    },
    async ({ html, options }) => {
      try {
        // Default options
        const defaultOptions = {
          wordwrap: 100,
          preserveNewlines: true,
          tables: true,
          preserveHrefLinks: true,
          headingStyle: 'linebreak' as const,
          hideQuotedContent: false
        };
        
        // Apply options if provided
        const convertOptions = {
          ...defaultOptions,
          ...options
        };
        
        // Convert HTML to plain text
        let plainText = htmlToText(html, convertOptions);
        
        // If hideQuotedContent is enabled, extract only the main message
        if (convertOptions.hideQuotedContent) {
          const parts = plainText.split('\n---\n');
          if (parts.length > 1) {
            plainText = parts[0] + '\n\n[Prior quoted messages removed]';
          }
        }
        
        return {
          content: [{ 
            type: 'text', 
            text: plainText
          }]
        };
      } catch (error) {
        return {
          content: [{ 
            type: 'text', 
            text: `Error converting HTML to text: ${error}`
          }],
          isError: true
        };
      }
    }
  );

  // Connect to the transport and start the server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  
  console.log('MCP server running. Use Ctrl+C to exit.');
}

// Run the server
if (require.main === module) {
  main().catch(error => {
    console.error('Error running MCP server:', error);
    process.exit(1);
  });
}