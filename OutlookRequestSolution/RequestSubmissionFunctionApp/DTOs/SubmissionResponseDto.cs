using System;
using System.Collections.Generic;

namespace RequestSubmissionFunctionApp.DTOs
{
    /// <summary>
    /// Data Transfer Object representing the response sent back to the Outlook client.
    /// </summary>
    public class SubmissionResponseDto
    {
        /// <summary>
        /// Indicating if the submission workflow completed successfully.
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// The unique ID generated for this specific submission tracking.
        /// </summary>
        public Guid SubmissionId { get; set; }

        /// <summary>
        /// The web URL of the SharePoint folder containing the uploaded attachments.
        /// </summary>
        public string FolderUrl { get; set; } = string.Empty;

        /// <summary>
        /// A status or summary message.
        /// </summary>
        public string Message { get; set; } = string.Empty;

        /// <summary>
        /// Detailed results for each individual attachment upload.
        /// </summary>
        public List<FileUploadResultDto> FileUploadResults { get; set; } = new List<FileUploadResultDto>();

        /// <summary>
        /// A list of error messages (e.g., validation failures or general exceptions).
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();
    }
}
