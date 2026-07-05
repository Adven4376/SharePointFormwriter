namespace RequestSubmissionFunctionApp.Configuration
{
    /// <summary>
    /// Configuration options for connecting to and interacting with SharePoint Online.
    /// </summary>
    public class SharePointSettings
    {
        /// <summary>
        /// The Microsoft Entra Tenant ID.
        /// </summary>
        public string TenantId { get; set; } = string.Empty;

        /// <summary>
        /// The Client (Application) ID of the registered Entra Application.
        /// </summary>
        public string ClientId { get; set; } = string.Empty;

        /// <summary>
        /// The Client Secret. Only populated locally or retrieved at runtime from Key Vault.
        /// </summary>
        public string ClientSecret { get; set; } = string.Empty;

        /// <summary>
        /// The SharePoint Site ID. Can be formatted as 'hostname,site-id,web-id'.
        /// </summary>
        public string SiteId { get; set; } = string.Empty;

        /// <summary>
        /// The ID or Name of the SharePoint Document Library where attachments are stored.
        /// </summary>
        public string LibraryId { get; set; } = string.Empty;

        /// <summary>
        /// The ID or Name of the target SharePoint List for submission metadata.
        /// </summary>
        public string ListId { get; set; } = string.Empty;

        /// <summary>
        /// The root folder path inside the Document Library for request attachments. Default is "RequestAttachments".
        /// </summary>
        public string RequestAttachmentsFolder { get; set; } = "RequestAttachments";
    }
}
