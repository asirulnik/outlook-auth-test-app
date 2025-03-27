import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { MailService } from './mailService';
import { htmlToText } from './htmlToText';

/**
 * This is an outline for the future MCP server implementation
 * It defines the expected tools and resources that will be exposed
 */

async function main() {
  // Create the MCP server
  const server = new McpServer({
    name: 'Outlook MCP Server',
    version: '1.0.0'
  });

  // Define the tools for the Outlook MCP server
  
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

  // 3. List emails tool
  server.tool(
    'list-emails',
    { 
      userEmail: z.string().email(),
      folderId: z.string(),
      limit: z.number().min(1).max(100).optional().default(25)
    },
    async ({ userEmail, folderId, limit }) => {
      try {
        const mailService = new MailService();
        const emails = await mailService.listEmails(folderId, userEmail, limit);
        
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

  // 4. Read email tool
  server.tool(
    'read-email',
    { 
      userEmail: z.string().email(),
      emailId: z.string(),
      convertHtmlToText: z.boolean().optional().default(true),
      htmlToTextOptions: z.object({
        wordwrap: z.union([z.number(), z.boolean()]).optional(),
        preserveNewlines: z.boolean().optional(),
        tables: z.boolean().optional(),
        preserveHrefLinks: z.boolean().optional(),
        headingStyle: z.enum(['underline', 'linebreak', 'hashify']).optional()
      }).optional()
    },
    async ({ userEmail, emailId, convertHtmlToText, htmlToTextOptions }) => {
      try {
        const mailService = new MailService();
        const email = await mailService.getEmail(emailId, userEmail);
        
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
        maxLineLength: z.number().optional()
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
          headingStyle: 'linebreak' as const
        };
        
        // Apply options if provided
        const convertOptions = {
          ...defaultOptions,
          ...options
        };
        
        // Convert HTML to plain text
        const plainText = htmlToText(html, convertOptions);
        
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
  // This is commented out as this is just an outline
  /*
  const transport = new StdioServerTransport();
  await server.connect(transport);
  */
  
  console.log('MCP server outline created');
}

// Only run this if executed directly
if (require.main === module) {
  main().catch(console.error);
}
