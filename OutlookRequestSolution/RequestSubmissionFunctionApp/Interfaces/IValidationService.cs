using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using RequestSubmissionFunctionApp.DTOs;

namespace RequestSubmissionFunctionApp.Interfaces
{
    /// <summary>
    /// Service responsible for validating submission requests on the server.
    /// </summary>
    public interface IValidationService
    {
        /// <summary>
        /// Validates a submission request, verifying fields, file counts, sizes, extensions, and content types.
        /// </summary>
        Task<(bool IsValid, List<string> Errors)> ValidateRequest(SubmissionRequestDto request);

        /// <summary>
        /// Validates the name field.
        /// </summary>
        Task<(bool IsValid, string? Error)> ValidateName(string name);

        /// <summary>
        /// Validates the email address.
        /// </summary>
        Task<(bool IsValid, string? Error)> ValidateEmail(string email);

        /// <summary>
        /// Validates the department field.
        /// </summary>
        Task<(bool IsValid, string? Error)> ValidateDepartment(string department);

        /// <summary>
        /// Validates the collection of file attachments.
        /// </summary>
        Task<(bool IsValid, List<string> Errors)> ValidateAttachments(List<IFormFile> attachments);

        /// <summary>
        /// Validates an individual file extension.
        /// </summary>
        Task<(bool IsValid, string? Error)> ValidateFileExtension(string fileName);

        /// <summary>
        /// Validates an individual file size.
        /// </summary>
        Task<(bool IsValid, string? Error)> ValidateFileSize(long fileSize);
    }
}
