using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using RequestSubmissionFunctionApp.Configuration;
using RequestSubmissionFunctionApp.DTOs;
using RequestSubmissionFunctionApp.Interfaces;
using RequestSubmissionFunctionApp.Utilities;

namespace RequestSubmissionFunctionApp.Services
{
    /// <summary>
    /// Service that implements verification rules for requests and attachments.
    /// </summary>
    public class ValidationService : IValidationService
    {
        private readonly ApplicationSettings _appSettings;

        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationService"/> class.
        /// </summary>
        /// <param name="appSettings">Injected application settings.</param>
        public ValidationService(ApplicationSettings appSettings)
        {
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
        }

        /// <inheritdoc/>
        public async Task<(bool IsValid, List<string> Errors)> ValidateRequest(SubmissionRequestDto request)
        {
            var errors = new List<string>();

            if (request == null)
            {
                errors.Add("Request payload cannot be empty.");
                return (false, errors);
            }

            // Validate Name
            var (isNameValid, nameError) = await ValidateName(request.Name);
            if (!isNameValid && nameError != null) errors.Add(nameError);

            // Validate Email
            var (isEmailValid, emailError) = await ValidateEmail(request.Email);
            if (!isEmailValid && emailError != null) errors.Add(emailError);

            // Validate Department
            var (isDeptValid, deptError) = await ValidateDepartment(request.Department);
            if (!isDeptValid && deptError != null) errors.Add(deptError);

            // Validate Description
            if (string.IsNullOrWhiteSpace(request.Description))
            {
                errors.Add("Description is required.");
            }

            // Validate Attachments
            var (areAttachmentsValid, attachmentErrors) = await ValidateAttachments(request.Attachments);
            if (!areAttachmentsValid)
            {
                errors.AddRange(attachmentErrors);
            }

            return (errors.Count == 0, errors);
        }

        /// <inheritdoc/>
        public Task<(bool IsValid, string? Error)> ValidateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return Task.FromResult<(bool, string?)>((false, "Name is required."));
            }
            return Task.FromResult<(bool, string?)>((true, null));
        }

        /// <inheritdoc/>
        public Task<(bool IsValid, string? Error)> ValidateEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return Task.FromResult<(bool, string?)>((false, "Email is required."));
            }

            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                if (addr.Address == email)
                {
                    return Task.FromResult<(bool, string?)>((true, null));
                }
            }
            catch
            {
                // Fallthrough to error
            }

            return Task.FromResult<(bool, string?)>((false, "Email address format is invalid."));
        }

        /// <inheritdoc/>
        public Task<(bool IsValid, string? Error)> ValidateDepartment(string department)
        {
            if (string.IsNullOrWhiteSpace(department))
            {
                return Task.FromResult<(bool, string?)>((false, "Department is required."));
            }
            return Task.FromResult<(bool, string?)>((true, null));
        }

        /// <inheritdoc/>
        public async Task<(bool IsValid, List<string> Errors)> ValidateAttachments(List<IFormFile> attachments)
        {
            var errors = new List<string>();
            const int maxAttachmentCount = 5;

            if (attachments == null)
            {
                return (true, errors);
            }

            if (attachments.Count > maxAttachmentCount)
            {
                errors.Add($"Maximum allowed attachments is {maxAttachmentCount}. Your request has {attachments.Count}.");
            }

            foreach (var attachment in attachments)
            {
                if (string.IsNullOrWhiteSpace(attachment.FileName))
                {
                    errors.Add("Attachment name cannot be empty.");
                    continue;
                }

                // Extension validation
                var (isExtValid, extError) = await ValidateFileExtension(attachment.FileName);
                if (!isExtValid && extError != null)
                {
                    errors.Add(extError);
                }

                // File size validation
                var (isSizeValid, sizeError) = await ValidateFileSize(attachment.Length);
                if (!isSizeValid && sizeError != null)
                {
                    errors.Add($"File '{attachment.FileName}': {sizeError}");
                }
            }

            return (errors.Count == 0, errors);
        }

        /// <inheritdoc/>
        public Task<(bool IsValid, string? Error)> ValidateFileExtension(string fileName)
        {
            var ext = FileUtility.GetExtension(fileName);
            
            // Check against whitelisted extensions from ApplicationSettings
            if (string.IsNullOrEmpty(ext) || !_appSettings.AllowedFileExtensions.Contains(ext))
            {
                return Task.FromResult<(bool, string?)>((
                    false, 
                    $"File '{fileName}' has an unsupported extension. Allowed formats are: {string.Join(", ", _appSettings.AllowedFileExtensions)}."
                ));
            }

            return Task.FromResult<(bool, string?)>((true, null));
        }

        /// <inheritdoc/>
        public Task<(bool IsValid, string? Error)> ValidateFileSize(long fileSize)
        {
            if (fileSize <= 0)
            {
                return Task.FromResult<(bool, string?)>((false, "File is empty."));
            }

            if (fileSize > _appSettings.MaximumFileSize)
            {
                var displayMax = FileUtility.ConvertSize(_appSettings.MaximumFileSize);
                return Task.FromResult<(bool, string?)>((
                    false, 
                    $"File size exceeds the maximum limit of {displayMax}."
                ));
            }

            return Task.FromResult<(bool, string?)>((true, null));
        }
    }
}
