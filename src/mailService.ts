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
}
