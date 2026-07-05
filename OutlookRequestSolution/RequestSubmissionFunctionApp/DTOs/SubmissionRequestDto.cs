using System.Collections.Generic;
using Microsoft.AspNetCore.Http;

namespace RequestSubmissionFunctionApp.DTOs
{
    /// <summary>
    /// Data Transfer Object representing the incoming submission request from the Outlook Add-in.
    /// </summary>
    public class SubmissionRequestDto
    {
        /// <summary>
        /// The name of the user submitting the request.
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// The email address of the user.
        /// </summary>
        public string Email { get; set; } = string.Empty;

        /// <summary>
        /// The department of the user.
        /// </summary>
        public string Department { get; set; } = string.Empty;

        /// <summary>
        /// The detailed description of the request.
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// The collection of attachments uploaded with the request, matching standard IFormFile.
        /// </summary>
        public List<IFormFile> Attachments { get; set; } = new List<IFormFile>();
    }
}
