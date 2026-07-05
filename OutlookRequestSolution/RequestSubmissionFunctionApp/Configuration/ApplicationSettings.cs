using System.Collections.Generic;

namespace RequestSubmissionFunctionApp.Configuration
{
    /// <summary>
    /// Configuration options for application-wide policies and message templates.
    /// </summary>
    public class ApplicationSettings
    {
        /// <summary>
        /// List of supported file extensions (including the dot).
        /// </summary>
        public List<string> AllowedFileExtensions { get; set; } = new List<string>
        {
            ".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".png", ".jpg", ".jpeg", ".txt", ".zip"
        };

        /// <summary>
        /// The maximum allowed file size in bytes for a single attachment. Default is 10MB.
        /// </summary>
        public long MaximumFileSize { get; set; } = 10 * 1024 * 1024;

        /// <summary>
        /// Name of the folder prefix to prepend to submission folders. Default is "Submission_".
        /// </summary>
        public string FolderPrefix { get; set; } = "Request_";

        /// <summary>
        /// Configurable success messages.
        /// </summary>
        public Dictionary<string, string> SuccessMessages { get; set; } = new Dictionary<string, string>
        {
            { "Default", "Your request has been successfully submitted and stored." }
        };

        /// <summary>
        /// Configurable error messages.
        /// </summary>
        public Dictionary<string, string> ErrorMessages { get; set; } = new Dictionary<string, string>
        {
            { "ValidationFailed", "Validation failed. Please correct the errors and try again." },
            { "UploadFailed", "File upload failed. Entire submission rolled back." },
            { "GenericError", "An error occurred while processing your request. Please contact support." }
        };
    }
}
