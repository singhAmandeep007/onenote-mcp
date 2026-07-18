#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import { DeviceCodeCredential } from '@azure/identity';
import fetch from 'node-fetch';

// Load environment variables
dotenv.config();

// Get the current file's directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

// Strip an optional "Bearer " prefix and surrounding whitespace from a token.
function normalizeAccessToken(value) {
  if (!value) return null;
  let token = String(value).trim();
  if (token.toLowerCase().startsWith('bearer ')) {
    token = token.slice(7).trim();
  }
  return token;
}

// A valid Microsoft Graph access token is a JWT: three base64url parts split by '.'.
// Personal-account tokens minted with the wrong scopes come back as a compact
// (non-JWT) string that Graph rejects with 401/40001.
function isLikelyJwt(token) {
  return typeof token === 'string' && token.split('.').length === 3;
}

// Create the MCP server
const server = new McpServer(
  { 
    name: "onenote",
    version: "1.0.0",
    description: "OneNote MCP Server" 
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

// Try to read the stored access token
let accessToken = null;
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error('Error reading access token file:', error.message);
}

// Alternatively, check if token is in environment variables
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

let graphClient = null;

// Client ID for Microsoft Graph API access
const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Explorer client ID

// 'common' accepts BOTH personal Microsoft accounts (MSA) and work/school (Azure AD).
const tenantId = process.env.GRAPH_TENANT || 'common';

// Resource-qualified, NON-".All" delegated scopes. Personal accounts cannot consent
// to ".All" scopes, which yields a token Graph rejects (401/40001). These work for all
// account types and cover reading/writing the signed-in user's own OneNote content.
const scopes = [
  'https://graph.microsoft.com/Notes.Read',
  'https://graph.microsoft.com/Notes.ReadWrite',
  'https://graph.microsoft.com/User.Read'
];

// Function to ensure Graph client is created
async function ensureGraphClient() {
  if (!graphClient) {
    // Read token from file if it exists
    try {
      if (fs.existsSync(tokenFilePath)) {
        const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
        try {
          // Try to parse as JSON first (new format)
          const parsedToken = JSON.parse(tokenData);
          accessToken = parsedToken.token;
        } catch (parseError) {
          // Fall back to using the raw token (old format)
          accessToken = tokenData;
        }
        accessToken = normalizeAccessToken(accessToken);
      }
    } catch (error) {
      console.error("Error reading token file:", error);
    }

    if (!accessToken) {
      throw new Error("Access token not found. Please authenticate first (run 'npm run auth' or use the 'authenticate' tool).");
    }

    // Personal Microsoft accounts return compact (non-JWT) Graph tokens — that's normal
    // and valid. Only warn; let Microsoft Graph be the authority on token validity.
    if (!isLikelyJwt(accessToken)) {
      console.error("Note: access token is a compact (non-JWT) token — normal for personal Microsoft accounts. Proceeding.");
    }

    // Create Microsoft Graph client (middleware-based auth provider)
    graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => accessToken
      }
    });
  }
  return graphClient;
}

// Create graph client with device code auth or access token
async function createGraphClient() {
  if (accessToken) {
    // Use access token if available
    graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          return accessToken;
        }
      }
    });
    return { type: 'token', client: graphClient };
  } else {
    // Use device code flow
    const credential = new DeviceCodeCredential({
      clientId: clientId,
      tenantId: tenantId,
      userPromptCallback: (info) => {
        // This will be shown to the user with the URL and code
        console.error('\n' + info.message);
      }
    });

    try {
      // Get an access token using device code flow
      const tokenResponse = await credential.getToken(scopes);
      
      // Save the token for future use
      accessToken = tokenResponse.token;
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));
      
      // Initialize Graph client with the token
      graphClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => {
            return accessToken;
          }
        }
      });
      
      return { type: 'device_code', client: graphClient };
    } catch (error) {
      console.error('Authentication error:', error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
}

// Tool for starting authentication flow
server.tool(
  "authenticate",
  "Start the authentication flow with Microsoft Graph",
  async () => {
    try {
      const result = await createGraphClient();
      if (result.type === 'device_code') {
        return { 
          content: [
            {
              type: "text",
              text: "Authentication started. Please check the console for the URL and code."
            }
          ]
        };
      } else {
        return { 
          content: [
            {
              type: "text",
              text: "Already authenticated with an access token."
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in authentication:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
);

// Tool for saving an access token provided by the user
server.tool(
  "saveAccessToken",
  "Save a Microsoft Graph access token for later use",
  async (params) => {
    try {
      // Save the token for future use
      const candidate = normalizeAccessToken(params.random_string);
      if (!candidate) {
        throw new Error("No token provided. Use the 'authenticate' tool or 'npm run auth' to obtain one.");
      }
      accessToken = candidate;
      const tokenData = JSON.stringify({ token: accessToken });
      fs.writeFileSync(tokenFilePath, tokenData);
      await createGraphClient();
      return { 
        content: [
          {
            type: "text",
            text: "Access token saved successfully"
          }
        ]
      };
    } catch (error) {
      console.error("Error saving access token:", error);
      throw new Error(`Failed to save access token: ${error.message}`);
    }
  }
);

// Tool for listing all notebooks
server.tool(
  "listNotebooks",
  "List all OneNote notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      // Return content as an array of text items
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing notebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

// Tool for getting notebook details
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/notebooks`).get();
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value[0])
          }
        ]
      };
    } catch (error) {
      console.error("Error getting notebook:", error);
      throw new Error(`Failed to get notebook: ${error.message}`);
    }
  }
);

// Tool for listing sections in a notebook
server.tool(
  "listSections",
  "List all sections in a notebook",
  async (params) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/sections`).get();
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing sections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
      
      if (sectionsResponse.value.length === 0) {
        return { 
          content: [
            {
              type: "text",
              text: "[]"
            }
          ]
        };
      }
      
      // Use the first section
      const sectionId = sectionsResponse.value[0].id;
      const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();
      
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing pages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page",
  async (params) => {
    try {
      console.error("GetPage called with params:", params);
      await ensureGraphClient();
      
      // First, list all pages to find the one we want
      const pagesResponse = await graphClient.api('/me/onenote/pages').get();
      console.error("Got", pagesResponse.value.length, "pages");
      
      let targetPage;
      
      // If a page ID is provided, use it to find the page
      if (params.random_string && params.random_string.length > 0) {
        const pageId = params.random_string;
        console.error("Looking for page with ID:", pageId);
        
        // Look for exact match first
        targetPage = pagesResponse.value.find(p => p.id === pageId);
        
        // If no exact match, try matching by title
        if (!targetPage) {
          console.error("No exact match, trying title search");
          targetPage = pagesResponse.value.find(p => 
            p.title && p.title.toLowerCase().includes(params.random_string.toLowerCase())
          );
        }
        
        // If still no match, try partial ID match
        if (!targetPage) {
          console.error("No title match, trying partial ID match");
          targetPage = pagesResponse.value.find(p => 
            p.id.includes(pageId) || pageId.includes(p.id)
          );
        }
      } else {
        // If no ID provided, use the first page
        console.error("No ID provided, using first page");
        targetPage = pagesResponse.value[0];
      }
      
      if (!targetPage) {
        throw new Error("Page not found");
      }
      
      console.error("Target page found:", targetPage.title);
      console.error("Page ID:", targetPage.id);
      
      try {
        const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage.id}/content`;
        console.error("Fetching content from:", url);
        
        // Make direct HTTP request with fetch
        const response = await fetch(url, {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });
        
        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
        }
        
        const content = await response.text();
        console.error(`Content received! Length: ${content.length} characters`);
        
        // Return the raw HTML content
        return {
          content: [
            {
              type: "text",
              text: content
            }
          ]
        };
      } catch (error) {
        console.error("Error getting content:", error);
        
        // Return a simple error message
        return {
          content: [
            {
              type: "text",
              text: `Error retrieving page content: ${error.message}`
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in getPage:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error in getPage: ${error.message}`
          }
        ]
      };
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section",
  async (params) => {
    try {
      await ensureGraphClient();
      // Get sections first
      const sectionsResponse = await graphClient.api(`/me/onenote/sections`).get();
      
      if (sectionsResponse.value.length === 0) {
        throw new Error("No sections found");
      }
      
      // Use the first section
      const sectionId = sectionsResponse.value[0].id;
      
      // Create simple HTML content
      const simpleHtml = `
        <!DOCTYPE html>
        <html>
          <head>
            <title>New Page</title>
          </head>
          <body>
            <p>This is a new page created via the Microsoft Graph API</p>
          </body>
        </html>
      `;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(simpleHtml);
      
      return { 
        content: [
          {
            type: "text",
            text: JSON.stringify(response)
          }
        ]
      };
    } catch (error) {
      console.error("Error creating page:", error);
      throw new Error(`Failed to create page: ${error.message}`);
    }
  }
);

// Tool for searching pages
server.tool(
  "searchPages",
  "Search for pages across notebooks",
  async (params) => {
    try {
      await ensureGraphClient();
      
      // Get all pages
      const response = await graphClient.api(`/me/onenote/pages`).get();
      
      // If search string is provided, filter the results
      if (params.random_string && params.random_string.length > 0) {
        const searchTerm = params.random_string.toLowerCase();
        const filteredPages = response.value.filter(page => {
          // Search in title
          if (page.title && page.title.toLowerCase().includes(searchTerm)) {
            return true;
          }
          return false;
        });
        
        return { 
          content: [
            {
              type: "text",
              text: JSON.stringify(filteredPages)
            }
          ]
        };
      } else {
        // Return all pages if no search term
        return { 
          content: [
            {
              type: "text",
              text: JSON.stringify(response.value)
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error searching pages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

// Connect to stdio and start server
async function main() {
  try {
    // Connect to standard I/O
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('Server started successfully.');
    console.error('Use the "authenticate" tool to start the authentication flow,');
    console.error('or use "saveAccessToken" if you already have a token.');
    
    // Keep the process alive
    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main(); 