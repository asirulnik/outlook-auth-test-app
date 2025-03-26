import { Client } from '@microsoft/microsoft-graph-client';
import { getGraphClient } from './authHelper';

// Interface for mail folder
export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount: number;
  unreadItemCount?: number;
  totalItemCount?: number;
}

// Interface for email message
export interface EmailMessage {
  id: string;
  subject?: string;
  from?: {
    emailAddress: {
      name?: string;
      address: string;
    }
  };
  receivedDateTime?: string;
  bodyPreview?: string;
  hasAttachments?: boolean;
  isRead?: boolean;
}

// Interface for email details
export interface EmailDetails extends EmailMessage {
  body?: {
    contentType: string;
    content: string;
  };
  toRecipients?: {
    emailAddress: {
      name?: string;
      address: string;
    }
  }[];
  ccRecipients?: {
    emailAddress: {
      name?: string;
      address: string;
    }
  }[];
  attachments?: {
    id: string;
    name: string;
    contentType: string;
    size: number;
  }[];
}

// Interface for creating a new email draft
export interface NewEmailDraft {
  subject: string;
  body: {
    contentType: string; // 'Text' or 'HTML'
    content: string;
  };
  toRecipients: {
    emailAddress: {
      address: string;
      name?: string;
    }
  }[];
  ccRecipients?: {
    emailAddress: {
      address: string;
      name?: string;
    }
  }[];
  bccRecipients?: {
    emailAddress: {
      address: string;
      name?: string;
    }
  }[];
}

// Interface for creating a new folder
export interface NewMailFolder {
  displayName: string;
  isHidden?: boolean;
}

// Interface for API errors
interface GraphApiError {
  statusCode?: number;
  message?: string;
}

// Service class for mail operations
export class MailService {
  private client: Client;

  constructor() {
    this.client = getGraphClient();
  }

  /**
   * Get mail folders for a specific user
   * @param userEmail Email address of the user (required for app-only authentication)
   * @returns List of mail folders
   */
  async getMailFolders(userEmail: string): Promise<MailFolder[]> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }
      
      // Build the API endpoint for the specified user
      const endpoint = `/users/${userEmail}/mailFolders`;
      
      // Query parameters to include child folder count and more details
      const queryParams = '?$top=100&$select=id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount';
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(`${endpoint}${queryParams}`)
        .get();
      
      return response.value;
    } catch (error) {
      console.error('Error getting mail folders:', error);
      throw error;
    }
  }

  /**
   * Get child folders for a specific mail folder
   * @param folderIdOrWellKnownName Folder ID or wellKnownName (like 'inbox')
   * @param userEmail Email address of the user (required for app-only authentication)
   * @returns List of child mail folders
   */
  async getChildFolders(folderIdOrWellKnownName: string, userEmail: string): Promise<MailFolder[]> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/mailFolders/${folderIdOrWellKnownName}/childFolders`;
      
      // Query parameters to include more details
      const queryParams = '?$select=id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount';
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(`${endpoint}${queryParams}`)
        .get();
      
      return response.value;
    } catch (error) {
      console.error('Error getting child folders:', error);
      throw error;
    }
  }

  /**
   * List emails in a folder
   * @param folderIdOrWellKnownName Folder ID or wellKnownName (like 'inbox')
   * @param userEmail Email address of the user
   * @param limit Number of emails to retrieve (default: 25)
   * @returns List of email messages
   */
  async listEmails(folderIdOrWellKnownName: string, userEmail: string, limit: number = 25): Promise<EmailMessage[]> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/mailFolders/${folderIdOrWellKnownName}/messages`;
      
      // Query parameters for pagination and fields
      const queryParams = `?$top=${limit}&$select=id,subject,from,receivedDateTime,bodyPreview,hasAttachments,isRead&$orderby=receivedDateTime desc`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(`${endpoint}${queryParams}`)
        .get();
      
      return response.value;
    } catch (error) {
      console.error('Error listing emails:', error);
      throw error;
    }
  }

  /**
   * Get a specific email with details
   * @param emailId ID of the email to retrieve
   * @param userEmail Email address of the user
   * @returns Email details
   */
  async getEmail(emailId: string, userEmail: string): Promise<EmailDetails> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/messages/${emailId}`;
      
      // Query parameters to include body and attachments
      const queryParams = '?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,hasAttachments,isRead';
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(`${endpoint}${queryParams}`)
        .get();
      
      // If the email has attachments, get them
      if (response.hasAttachments) {
        const attachmentsEndpoint = `/users/${userEmail}/messages/${emailId}/attachments`;
        const attachmentsResponse = await this.client
          .api(attachmentsEndpoint)
          .get();
        
        response.attachments = attachmentsResponse.value;
      }
      
      return response;
    } catch (error) {
      console.error('Error getting email details:', error);
      throw error;
    }
  }

  /**
   * Move an email to another folder
   * @param emailId ID of the email to move
   * @param destinationFolderId ID of the destination folder
   * @param userEmail Email address of the user
   * @returns Moved email message
   */
  async moveEmail(emailId: string, destinationFolderId: string, userEmail: string): Promise<EmailMessage> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/messages/${emailId}/move`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .post({
          destinationId: destinationFolderId
        });
      
      return response;
    } catch (error) {
      console.error('Error moving email:', error);
      throw error;
    }
  }

  /**
   * Copy an email to another folder
   * @param emailId ID of the email to copy
   * @param destinationFolderId ID of the destination folder
   * @param userEmail Email address of the user
   * @returns Copied email message
   */
  async copyEmail(emailId: string, destinationFolderId: string, userEmail: string): Promise<EmailMessage> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/messages/${emailId}/copy`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .post({
          destinationId: destinationFolderId
        });
      
      return response;
    } catch (error) {
      console.error('Error copying email:', error);
      throw error;
    }
  }

  /**
   * Create a new draft email
   * @param draft Draft email content
   * @param userEmail Email address of the user
   * @returns Created draft email
   */
  async createDraft(draft: NewEmailDraft, userEmail: string): Promise<EmailMessage> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint for creating a message in drafts folder
      const endpoint = `/users/${userEmail}/messages`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .post(draft);
      
      return response;
    } catch (error) {
      console.error('Error creating draft:', error);
      throw error;
    }
  }

  /**
   * Create a new mail folder
   * @param newFolder New folder details
   * @param parentFolderId Optional parent folder ID (if not provided, creates at root)
   * @param userEmail Email address of the user
   * @returns Created mail folder
   */
  async createFolder(newFolder: NewMailFolder, userEmail: string, parentFolderId?: string): Promise<MailFolder> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      let endpoint = `/users/${userEmail}`;
      
      if (parentFolderId) {
        endpoint += `/mailFolders/${parentFolderId}/childFolders`;
      } else {
        endpoint += '/mailFolders';
      }
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .post(newFolder);
      
      return response;
    } catch (error) {
      console.error('Error creating folder:', error);
      throw error;
    }
  }

  /**
   * Update a mail folder's properties (rename)
   * @param folderId ID of the folder to update
   * @param updatedFolder Updated folder properties
   * @param userEmail Email address of the user
   * @returns Updated mail folder
   */
  async updateFolder(folderId: string, updatedFolder: Partial<NewMailFolder>, userEmail: string): Promise<MailFolder> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/mailFolders/${folderId}`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .patch(updatedFolder);
      
      return response;
    } catch (error) {
      console.error('Error updating folder:', error);
      throw error;
    }
  }

  /**
   * Move a folder to another parent folder
   * @param folderId ID of the folder to move
   * @param destinationParentFolderId ID of the destination parent folder
   * @param userEmail Email address of the user
   * @returns Moved mail folder
   */
  async moveFolder(folderId: string, destinationParentFolderId: string, userEmail: string): Promise<MailFolder> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/mailFolders/${folderId}/move`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .post({
          destinationId: destinationParentFolderId
        });
      
      return response;
    } catch (error) {
      console.error('Error moving folder:', error);
      throw error;
    }
  }

  /**
   * Copy a folder to another parent folder
   * Note: This might not be supported by the Microsoft Graph API
   * @param folderId ID of the folder to copy
   * @param destinationParentFolderId ID of the destination parent folder
   * @param userEmail Email address of the user
   * @returns Copied mail folder
   */
  async copyFolder(folderId: string, destinationParentFolderId: string, userEmail: string): Promise<MailFolder> {
    try {
      if (!userEmail) {
        throw new Error('User email is required for application permissions flow');
      }

      // Build the API endpoint
      const endpoint = `/users/${userEmail}/mailFolders/${folderId}/copy`;
      
      // Make the request to Microsoft Graph
      const response = await this.client
        .api(endpoint)
        .post({
          destinationId: destinationParentFolderId
        });
      
      return response;
    } catch (error) {
      const apiError = error as GraphApiError;
      // Check if the error is because the API doesn't support folder copying
      if (apiError.statusCode === 501 || apiError.message?.includes('Not Implemented')) {
        throw new Error('Folder copying is not supported by the Microsoft Graph API');
      }
      console.error('Error copying folder:', error);
      throw error;
    }
  }
}
