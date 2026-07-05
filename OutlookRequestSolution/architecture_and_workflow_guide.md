# Architectural & Workflow Guide: Outlook Request Submission System

This guide explains the entire system design, technology stack, security/authentication mechanisms, and step-by-step execution flow of the Request Submission application. It includes visible code snippets, file links, and detailed explanations of the logic behind each phase.

---

## 🌟 Section 1: Plain English Overview (How It Works)

Imagine a user working in Microsoft Outlook who wants to submit a support request or project files to the company's SharePoint site.

```text
[ Outlook User Interface ] 
      │ 
      ▼ (Clicks Submit)
[ Task Pane Add-In (Frontend) ] ────(HTTP POST: Form Fields + Attachments)───► [ Azure Function (Backend) ]
                                                                                       │
                                                                         ┌─────────────┴─────────────┐
                                                                         ▼                           ▼
                                                             [ SharePoint Library ]         [ SharePoint List ]
                                                             (Files in GUID Folder)         (Record & Link)
```

1.  **Filling out the Form**: The user opens the custom **Enterprise Submit Hub** add-in directly within Outlook. They fill in their Name, Email, Department, and Description, and select up to 5 attachments (like PDFs, Excel sheets, or images).
2.  **Sending the Data**: When they click **Submit**, the frontend bundle packages all text entries and attachments into a single HTTP request packet and sends it securely to our serverless backend (Azure Function).
3.  **Processing & Storage (The "Double Save")**:
    *   **The Folder**: The backend generates a unique identification code (a **GUID**, like `a1b2c3d4-e5f6...`). It tells SharePoint to create a folder with this ID in the Document Library and uploads all files inside it.
    *   **The List**: Once the files are uploaded, SharePoint returns the web link to that folder. The backend then records the user's details (Name, Email, Description) along with the folder link as a single entry (row) in a SharePoint List.
4.  **Completion**: The backend returns a success message to the frontend, which displays a confirmation to the user. If any step fails, the backend cleans up after itself by deleting the folder so no orphaned files are left behind.

---

## ⚙️ Section 2: Core Configuration & Secret Keys

To run the application, several key identifiers must be declared. These establish the credentials and locations of resources inside Microsoft 365.

### 1. Where Configuration Keys are Declared
*   **Locally**: Configured inside **[local.settings.json](file:///e:/aditya/azure/OutlookRequestSolution/RequestSubmissionFunctionApp/local.settings.json)**:
    ```json
    {
      "IsEncrypted": false,
      "Values": {
        "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
        "AzureWebJobsStorage": "UseDevelopmentStorage=true",
        "SharePoint__TenantId": "your-tenant-id",
        "SharePoint__ClientId": "your-client-id",
        "SharePoint__ClientSecret": "your-client-secret-only-for-local",
        "SharePoint__SiteId": "your-sharepoint-site-id",
        "SharePoint__LibraryId": "your-document-library-id-or-name",
        "SharePoint__ListId": "your-sharepoint-list-id-or-name",
        "SharePoint__RequestAttachmentsFolder": "RequestAttachments"
      }
    }
    ```
*   **In Production (Azure)**: Configured as Environment Variables in the Azure Function App's Application Settings.

### 2. Purpose of the Key Identifiers
*   **TenantId**: Identifies your company's unique Microsoft 365 cloud instance. It acts as the home directory of your accounts.
*   **ClientId**: The unique Application ID of your registered app in Microsoft Entra ID. It tells Microsoft *which* program is requesting access.
*   **ClientSecret**: The password for your Application. It verifies that your program is legitimate and authorized to request data.
*   **SiteId**: The unique ID of your SharePoint Site. It tells the SDK *which* site collection to target.
*   **LibraryId**: The identifier of the document library (drive) hosting attachments.
*   **ListId**: The identifier of the SharePoint List storing user requests metadata.

---

## 🔑 Section 3: Microsoft Graph API & The Singleton Client

The **Microsoft Graph API** is the unified endpoint to interact with SharePoint, Outlook, and Microsoft 365 data. 

### 1. Token Acquisition Flow
To connect to SharePoint, the Azure Function App must acquire an Access Token from Microsoft Entra ID.
1. The Function presents its `ClientId`, `TenantId`, and `ClientSecret` to the Entra ID authorization server.
2. Entra ID validates the client secret and generates a cryptographically signed OAuth JSON Web Token (JWT).
3. The Function app appends this token as a Bearer string in the headers of all subsequent Graph API HTTP requests.

### 2. Graph Client Singleton Registration in [Program.cs](file:///e:/aditya/azure/OutlookRequestSolution/RequestSubmissionFunctionApp/Program.cs)
Rather than instantiating the credentials client and HTTP adapter on every incoming request, the client is registered as a **Singleton** (single instance shared across the entire application lifecycle). This prevents socket exhaustion and boosts performance.

```csharp
// Register GraphServiceClient with secure Authentication
services.AddSingleton<GraphServiceClient>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<Program>>();
    
    // Read authentication parameters from settings
    string tenantId = sharePointSettings.TenantId;
    string clientId = sharePointSettings.ClientId;
    string clientSecret = sharePointSettings.ClientSecret;
    string keyVaultUrl = configuration["AzureAd__KeyVaultUrl"] ?? string.Empty;

    // Secure fallback: Retrieve secret from Key Vault if missing locally
    if (string.IsNullOrEmpty(clientSecret) && !string.IsNullOrEmpty(keyVaultUrl))
    {
        try
        {
            logger.LogInformation("Retrieving SharePoint Client Secret from Azure Key Vault: {VaultUrl}", keyVaultUrl);
            var secretClient = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());
            KeyVaultSecret secret = secretClient.GetSecret("SharePoint--ClientSecret");
            clientSecret = secret.Value;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to retrieve Client Secret from Azure Key Vault.");
        }
    }

    // Initialize using Client Secret Credentials if keys are present
    if (!string.IsNullOrEmpty(tenantId) && !string.IsNullOrEmpty(clientId) && !string.IsNullOrEmpty(clientSecret))
    {
        logger.LogInformation("Initializing GraphServiceClient using ClientSecretCredential.");
        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        return new GraphServiceClient(credential);
    }
    else
    {
        logger.LogInformation("Initializing GraphServiceClient using DefaultAzureCredential (e.g. Managed Identity).");
        var credential = new DefaultAzureCredential();
        return new GraphServiceClient(credential);
    }
});
```

---

## ⚡ Section 4: The Client-Side HTTP Trigger (Frontend to Backend)

Here is how data is packaged in the frontend and captured by the backend Function.

### 1. Frontend: Dispatching the Request in [api.service.js](file:///e:/aditya/azure/OutlookRequestSolution/OutlookAddin/services/api.service.js)
The frontend client uses standard HTML5 `FormData` to bundle form variables and multiple binary file objects:

```javascript
const apiService = {
    apiEndpoint: 'http://localhost:7071/api/submit',

    async submitRequest(fields, files) {
        const formData = new FormData();

        // 1. Append text fields
        formData.append('Name', fields.name.trim());
        formData.append('Email', fields.email.trim());
        formData.append('Department', fields.department.trim());
        formData.append('Description', fields.description.trim());

        // 2. Append attachment files
        if (files && files.length > 0) {
            for (let i = 0; i < files.length; i++) {
                formData.append('Attachments', files[i], files[i].name);
            }
        }

        // 3. Make HTTP POST Request
        const response = await fetch(this.apiEndpoint, {
            method: 'POST',
            body: formData // Content-Type: multipart/form-data with boundary
        });

        const result = await response.json();
        if (!response.ok) {
            throw new Error(result.message || 'Request failed');
        }
        return result;
    }
};
```

### 2. Backend: Capturing the HTTP POST in [SubmitRequestFunction.cs](file:///e:/aditya/azure/OutlookRequestSolution/RequestSubmissionFunctionApp/Functions/SubmitRequestFunction.cs)
The backend Azure Function captures this payload using an `HttpTrigger`:

```csharp
[Function("SubmitRequest")]
public async Task<HttpResponseData> Run(
    [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "submit")] HttpRequestData req)
{
    _logger.LogInformation("HTTP Trigger 'SubmitRequest' received a request.");

    // Validate Content-Type
    if (!req.Headers.TryGetValues("Content-Type", out var contentTypeHeaders))
    {
        var errorResponse = req.CreateResponse(HttpStatusCode.BadRequest);
        await errorResponse.WriteAsJsonAsync(new { Error = "Content-Type header is required." });
        return errorResponse;
    }

    var contentType = contentTypeHeaders.FirstOrDefault() ?? string.Empty;
    if (!contentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
    {
        var errorResponse = req.CreateResponse(HttpStatusCode.BadRequest);
        await errorResponse.WriteAsJsonAsync(new { Error = "Content-Type must be multipart/form-data." });
        return errorResponse;
    }

    try
    {
        // Parse request stream with security size limit check
        SubmissionRequestDto requestDto = await MultipartParser.ParseAsync(req.Body, contentType, _appSettings.MaximumFileSize);

        // Orchestrate submission flow
        SubmissionResponseDto responseDto = await _submissionService.ProcessSubmission(requestDto);

        var httpStatusCode = responseDto.Success ? HttpStatusCode.OK : HttpStatusCode.BadRequest;
        var response = req.CreateResponse(httpStatusCode);
        await response.WriteAsJsonAsync(responseDto);
        return response;
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "Unexpected error occurred during execution.");
        var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
        await errorResponse.WriteAsJsonAsync(new { Error = "An unexpected error occurred." });
        return errorResponse;
    }
}
```

---

## 📂 Section 5: The File Storage Pipeline (SharePoint Library)

The orchestrator service generates a `Guid` which acts as the target folder name inside the Document Library.

### 1. Orchestrating Folder Creation in [SubmissionService.cs](file:///e:/aditya/azure/OutlookRequestSolution/RequestSubmissionFunctionApp/Services/SubmissionService.cs)
```csharp
// Generates unique GUID tracking ID
var submissionId = Guid.NewGuid(); // e.g. 5d5a7d8e-c90a-41e9-9a2c-d6a5d4f3b2c1

// Create the target folder in SharePoint Document Library
string folderUrl = await _documentService.CreateSubmissionFolder(submissionId);

// Upload each binary attachment
if (request.Attachments != null && request.Attachments.Count > 0)
{
    fileUploadResults = await _documentService.UploadMultipleFiles(submissionId, request.Attachments);
}
```

### 2. Creating the Folder in [SharePointDocumentService.cs](file:///e:/aditya/azure/OutlookRequestSolution/RequestSubmissionFunctionApp/Services/SharePointDocumentService.cs)
The Document Service resolves the parent library and posts a new `DriveItem` folder structure:

```csharp
public async Task<string> CreateSubmissionFolder(Guid submissionId)
{
    var drive = await GetDriveAsync(); // Uses memory caching to prevent listing drives repeatedly
    var submissionFolderName = submissionId.ToString();

    var newFolder = new DriveItem
    {
        Name = submissionFolderName,
        Folder = new Folder(),
        AdditionalData = new Dictionary<string, object>
        {
            { "@microsoft.graph.conflictBehavior", "replace" }
        }
    };

    // Post to Microsoft Graph API
    var createdFolder = await _graphServiceClient.Drives[drive.Id]
        .Root
        .ItemWithPath(_settings.RequestAttachmentsFolder)
        .Children
        .PostAsync(newFolder);

    return createdFolder.WebUrl; // Returns the direct SharePoint Web URL to the folder
}
```

---

## 📋 Section 6: The Metadata & Link Storage (SharePoint List)

Once the files are successfully uploaded, we save the user's form details alongside the generated folder web link (`folderUrl`) to the SharePoint List. 

### 1. Creating the List Item in [SharePointListService.cs](file:///e:/aditya/azure/OutlookRequestSolution/RequestSubmissionFunctionApp/Services/SharePointListService.cs)
The backend issues a POST request to Graph API list items:

```csharp
public async Task<string> CreateSubmissionItem(Guid submissionId, SubmissionRequestDto request, string folderUrl)
{
    var listItem = new ListItem
    {
        Fields = new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object>
            {
                { "Title", submissionId.ToString() }, // Standard identifier field
                { "SubmissionId", submissionId.ToString() },
                { "Name", request.Name },
                { "Email", request.Email },
                { "Department", request.Department },
                { "Description", request.Description },
                { "FolderURL", folderUrl }, // Stores the hyperlink to the uploaded files folder
                { "CreatedDate", DateTimeOffset.UtcNow.ToString("o") }
            }
        }
    };

    var createdItem = await _graphServiceClient.Sites[_settings.SiteId]
        .Lists[_settings.ListId]
        .Items
        .PostAsync(listItem);

    return createdItem.Id;
}
```

### 2. How the Clickable Folder Link Opens Files
*   The `FolderURL` field populated in the list is the official **SharePoint Web URL** (e.g. `https://yourcompany.sharepoint.com/sites/yoursite/RequestAttachments/5d5a7d8e-c90a-41e9-9a2c-d6a5d4f3b2c1`).
*   When administrators or team members view the SharePoint List inside their web browser, they can click on the `FolderURL` hyperlink. 
*   SharePoint natively redirects their browser window directly into the Document Library interface, showing the exact folder containing the user's uploaded attachment files.

---

## ⚠️ Section 7: Key Vault Transition (Local vs. Production Secrets)

Storing client secrets in text config files poses a high security risk. We have implemented a secure pipeline that automatically transitions between local config and Azure Key Vault in production.

### Detailed Migration Workflow:

```text
[ local.settings.json ]
  "SharePoint__ClientSecret"
         │
         ├──► (Local Dev): Program.cs binds and uses Client Secret directly.
         │
[ Azure Key Vault ]
  "SharePoint--ClientSecret" (Secure Vault)
         │
         ├──► (Production): Set "AzureAd__KeyVaultUrl" environment variable.
         └──► Program.cs detects setting, fetches secret via Managed Identity, and overrides value.
```

1.  **Configure Azure Key Vault**:
    Deploy a Key Vault in your subscription and add a secret named `SharePoint--ClientSecret`.
2.  **Managed Identity access**:
    Enable System-Assigned Managed Identity on the Azure Function App, and assign a Key Vault access policy granting the Function Identity `Get` permissions on Secrets.
3.  **Application Settings Configuration**:
    In your Azure Function settings, add:
    *   `AzureAd__KeyVaultUrl` = `https://<YOUR_VAULT_NAME>.vault.azure.net/`
    *   Leave `SharePoint__ClientSecret` **empty or deleted** in your Azure App settings.
4.  **Runtime Retrieval Logic**:
    At startup, `Program.cs` sees that `clientSecret` is empty but a Key Vault URL is supplied, calling:
    ```csharp
    var secretClient = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());
    KeyVaultSecret secret = secretClient.GetSecret("SharePoint--ClientSecret");
    clientSecret = secret.Value;
    ```
    This injects the credential securely into the singleton client without storing credentials in files.
