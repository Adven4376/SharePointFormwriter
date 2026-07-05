namespace RequestSubmissionFunctionApp.DTOs
{
    /// <summary>
    /// DTO representing the result of an individual file upload operation to SharePoint.
    /// </summary>
    public class FileUploadResultDto
    {
        /// <summary>
        /// The name of the file uploaded.
        /// </summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>
        /// The absolute URL or path to the file in SharePoint.
        /// </summary>
        public string FileUrl { get; set; } = string.Empty;

        /// <summary>
        /// Indicating whether the file was uploaded successfully.
        /// </summary>
        public bool UploadStatus { get; set; }

        /// <summary>
        /// The size of the uploaded file in bytes.
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// Error message if the upload failed.
        /// </summary>
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
