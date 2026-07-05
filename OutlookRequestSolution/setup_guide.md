# Setup & Requirements Guide: Outlook Request Submission System

This guide outlines all the resources, accounts, configurations, credentials, and API permissions required to configure and run the Outlook Add-in and Azure Function App backend.

---

## 📋 Table of Contents
1. [Required Accounts](#1-required-accounts)
2. [Entra ID (Azure AD) App Registration](#2-entra-id-azure-ad-app-registration)
3. [Microsoft Graph API Permissions](#3-microsoft-graph-api-permissions)
4. [SharePoint Setup & Identification](#4-sharepoint-setup--identification)
5. [Local Settings Configuration (`local.settings.json`)](#5-local-settings-configuration-localsettingsjson)
6. [Local Execution Steps](#6-local-execution-steps)

---

## 1. Required Accounts
To fully deploy and test the solution, you need:
*   **Azure Subscription**: To host the Azure Function App (and optional Key Vault to store secrets).
*   **Microsoft 365 Developer or Enterprise Account**: To access SharePoint Online, Outlook, and Microsoft Entra ID (formerly Azure AD).
    *   *Tip*: If you don't have one, you can sign up for a free sandbox tenant through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).

---

## 2. Entra ID (Azure AD) App Registration
The Azure Function App communicates with SharePoint Online using the Microsoft Graph API. It authenticates as a daemon service (Application Identity) using client credentials.

### Steps to Register the App:
1. Log in to the [Azure Portal](https://portal.azure.com/) or [Entra Admin Center](https://entra.microsoft.com/).
2. Navigate to **Microsoft Entra ID** $\rightarrow$ **App registrations** $\rightarrow$ **New registration**.
3. Configure the following:
   *   **Name**: `Request-Submission-Backend` (or similar).
   *   **Supported account types**: Accounts in this organizational directory only (Single tenant).
   *   **Redirect URI**: Leave empty.
4. Click **Register**.

### Retrieve IDs:
Once registered, copy the following values from the **Overview** pane:
*   **Application (client) ID** (Referred to as `ClientId`)
*   **Directory (tenant) ID** (Referred to as `TenantId`)

### Generate Client Secret:
1. In the left menu of the app registration, select **Certificates & secrets** $\rightarrow$ **Client secrets** $\rightarrow$ **New client secret**.
2. Add a description (e.g., `LocalDevSecret`) and choose an expiration timeline.
3. Click **Add**.
4. **IMPORTANT**: Copy the secret **Value** immediately. It will be hidden permanently once you navigate away from the page.

---

## 3. Microsoft Graph API Permissions
The application needs permissions to read/write documents in the SharePoint Document Library and list items in the SharePoint List.

1. Go to your App Registration in the Azure portal and select **API permissions** $\rightarrow$ **Add a permission**.
2. Select **Microsoft Graph**.
3. Choose **Application permissions** (since the backend runs as a background service without user login context).
4. Add the following permissions:
   *   `Sites.ReadWrite.All` - Allows the backend to create folders, upload attachments, and create/update list items in your SharePoint sites.
5. Click **Add permissions**.
6. **Grant Admin Consent**: Click the button **"Grant admin consent for [Your Tenant Name]"** (this step requires an Administrator account).

---

## 4. SharePoint Setup & Identification

You need a SharePoint Site containing a **Document Library** (for attachments) and a **List** (for metadata entries).

### A. Document Library Setup
*   Create a SharePoint Document Library or use an existing one (e.g., named `RequestAttachments`).
*   The folder hierarchy created at runtime will be: `{DocumentLibraryRoot}/RequestAttachments/{SubmissionGUID}/`.

### B. SharePoint List Setup
Create a SharePoint List (e.g., named `RequestSubmissions`) with the following custom columns (all created as "Single line of text" unless specified otherwise):
*   **Title** (Rename/use the default text field) $\rightarrow$ Stores the `SubmissionId` GUID.
*   **SubmissionId** (Single line of text)
*   **Name** (Single line of text)
*   **Email** (Single line of text)
*   **Department** (Single line of text)
*   **Description** (Multiple lines of text)
*   **FolderURL** (Hyperlink or text)
*   **CreatedDate** (Single line of text or Date/Time)

### C. Find SharePoint Identifiers
Microsoft Graph SDK needs specific IDs to point to your SharePoint resources.

#### 1. Retrieve Site ID
SharePoint Site IDs are formatted as: `hostname,site-collection-id,web-id`.
To retrieve your Site ID, perform an HTTP GET request to the following Graph API endpoint (you can use the [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)):
```http
GET https://graph.microsoft.com/v1.0/sites/{tenant-domain}.sharepoint.com:/sites/{your-site-name}
```
*Example Response:*
```json
{
    "id": "tenantname.sharepoint.com,a1b2c3d4-e5f6-7a8b-9c0d-1e2f3a4b5c6d,f9e8d7c6-b5a4-3f2e-1d0c-9b8a7f6e5d4c"
}
```
Copy the entire `id` string (e.g., `tenantname.sharepoint.com,a1b2c3...`). This is your `SiteId`.

#### 2. Retrieve Document Library ID
Get the ID of your Document Library using:
```http
GET https://graph.microsoft.com/v1.0/sites/{SiteId}/drives
```
Look for the drive matching the library name (e.g., `RequestAttachments`) and copy its `id`.

#### 3. Retrieve SharePoint List ID
Get the ID of your SharePoint List using:
```http
GET https://graph.microsoft.com/v1.0/sites/{SiteId}/lists
```
Look for the list matching your submissions list name and copy its `id`.

---

## 5. Local Settings Configuration (`local.settings.json`)

Fill in your retrieved IDs in the **local.settings.json** file:

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    
    "SharePoint__TenantId": "<YOUR_TENANT_ID>",
    "SharePoint__ClientId": "<YOUR_CLIENT_ID>",
    "SharePoint__ClientSecret": "<YOUR_CLIENT_SECRET>",
    "SharePoint__SiteId": "<YOUR_SHAREPOINT_SITE_ID>",
    "SharePoint__LibraryId": "<YOUR_DOCUMENT_LIBRARY_ID_OR_NAME>",
    "SharePoint__ListId": "<YOUR_SHAREPOINT_LIST_ID_OR_NAME>",
    "SharePoint__RequestAttachmentsFolder": "RequestAttachments",
    
    "Application__FolderPrefix": "Request_",
    "Application__MaximumFileSize": "10485760",
    
    "LibraryName": "RequestAttachments",
    "FunctionKey": "your-function-key-if-needed"
  },
  "Host": {
    "LocalHttpPort": 7071,
    "CORS": "*",
    "CORSCredentials": false
  }
}
```

---

## 6. Local Execution Steps

### Step 1: Run the Backend (Azure Function App)
Ensure you have the Azure Functions Core Tools installed.
Run the project using the local .NET binary:
1. Open PowerShell and navigate to the project directory:
   ```powershell
   cd e:\aditya\azure\OutlookRequestSolution\RequestSubmissionFunctionApp
   ```
2. Start the local Azure Functions runtime host:
   ```powershell
   func start
   ```
   *(Or build and run via your IDE or using the local dotnet runner: `..\..\.dotnet\dotnet.exe run`)*
3. Verify that the console output outputs the local endpoint:
   `SubmitRequest: [POST] http://localhost:7071/api/submit`

### Step 2: Run the Frontend (Outlook Add-in)
1. Navigate to the add-in directory:
   ```powershell
   cd e:\aditya\azure\OutlookRequestSolution\OutlookAddin
   ```
2. Install dependencies (if you use standard Node packages for serving/debugging):
   ```bash
   npm install
   ```
3. Start the local dev server (defaulting to port `3000` with HTTPS enabled):
   ```bash
   npm start
   ```
4. Sideload the `manifest.xml` file inside Outlook (Outlook on the Web or Desktop) to load the add-in.
