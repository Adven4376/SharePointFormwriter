using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using RequestSubmissionFunctionApp.Configuration;
using RequestSubmissionFunctionApp.DTOs;
using RequestSubmissionFunctionApp.Interfaces;

namespace RequestSubmissionFunctionApp.Services
{
    /// <summary>
    /// Service that coordinates the entire request submission workflow.
    /// </summary>
    public class SubmissionService : ISubmissionService
    {
        private readonly IValidationService _validationService;
        private readonly ISharePointDocumentService _documentService;
        private readonly ISharePointListService _listService;
        private readonly ApplicationSettings _appSettings;
        private readonly ILogger<SubmissionService> _logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubmissionService"/> class.
        /// </summary>
        public SubmissionService(
            IValidationService validationService,
            ISharePointDocumentService documentService,
            ISharePointListService listService,
            ApplicationSettings appSettings,
            ILogger<SubmissionService> logger)
        {
            _validationService = validationService ?? throw new ArgumentNullException(nameof(validationService));
            _documentService = documentService ?? throw new ArgumentNullException(nameof(documentService));
            _listService = listService ?? throw new ArgumentNullException(nameof(listService));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<SubmissionResponseDto> ProcessSubmission(SubmissionRequestDto request)
        {
            var submissionId = Guid.NewGuid();
            _logger.LogInformation("Starting submission workflow. Generated Submission ID: {SubmissionId}.", submissionId);

            // 1. Server-side Validation
            var (isValid, validationErrors) = await _validationService.ValidateRequest(request);
            if (!isValid)
            {
                _logger.LogWarning("Validation failed for submission {SubmissionId}. Errors: {Errors}", 
                    submissionId, string.Join("; ", validationErrors));
                
                var errorMsg = _appSettings.ErrorMessages.TryGetValue("ValidationFailed", out var valMsg) 
                    ? valMsg 
                    : "Validation failed.";

                return new SubmissionResponseDto
                {
                    Success = false,
                    SubmissionId = submissionId,
                    Message = errorMsg,
                    Errors = validationErrors
                };
            }

            string folderUrl = string.Empty;
            var fileUploadResults = new List<FileUploadResultDto>();

            try
            {
                // 2. Create Unique folder in SharePoint Document Library
                folderUrl = await _documentService.CreateSubmissionFolder(submissionId);

                // 3. Upload Attachments (if any)
                if (request.Attachments != null && request.Attachments.Count > 0)
                {
                    fileUploadResults = await _documentService.UploadMultipleFiles(submissionId, request.Attachments);

                    // Check if any individual file upload failed
                    var failedUploads = fileUploadResults.Where(r => !r.UploadStatus).ToList();
                    if (failedUploads.Any())
                    {
                        var failedFileNames = string.Join(", ", failedUploads.Select(f => f.FileName));
                        _logger.LogError("File upload failed for attachments: {FileNames}. Triggering rollback.", failedFileNames);

                        // Rollback: delete the created SharePoint folder
                        await _documentService.DeleteFolder(submissionId);

                        var uploadErrorMsg = _appSettings.ErrorMessages.TryGetValue("UploadFailed", out var upMsg) 
                            ? upMsg 
                            : "Upload failed.";

                        return new SubmissionResponseDto
                        {
                            Success = false,
                            SubmissionId = submissionId,
                            Message = $"{uploadErrorMsg} (Failed files: {failedFileNames})",
                            FileUploadResults = fileUploadResults,
                            Errors = failedUploads.Select(f => $"{f.FileName}: {f.ErrorMessage}").ToList()
                        };
                    }
                }

                // 4. Create SharePoint List Item (metadata storage)
                _logger.LogInformation("Uploading metadata for submission {SubmissionId} to SharePoint List.", submissionId);
                
                try
                {
                    await _listService.CreateSubmissionItem(submissionId, request, folderUrl);
                }
                catch (Exception listEx)
                {
                    _logger.LogError(listEx, "Failed to create SharePoint List metadata entry. Triggering folder rollback.");
                    
                    // Rollback: delete the uploaded attachments folder so we don't leave orphaned files
                    await _documentService.DeleteFolder(submissionId);
                    
                    throw; // Re-throw to return structured error response in outer catch
                }

                _logger.LogInformation("Submission workflow completed successfully for {SubmissionId}.", submissionId);

                var successMsg = _appSettings.SuccessMessages.TryGetValue("Default", out var sMsg) 
                    ? sMsg 
                    : "Submission completed successfully.";

                return new SubmissionResponseDto
                {
                    Success = true,
                    SubmissionId = submissionId,
                    Message = successMsg,
                    FolderUrl = folderUrl,
                    FileUploadResults = fileUploadResults
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred during submission workflow for {SubmissionId}.", submissionId);
                
                var genericErrorMsg = _appSettings.ErrorMessages.TryGetValue("GenericError", out var genMsg) 
                    ? genMsg 
                    : "An internal server error occurred.";

                // Security: Do not expose internal details in production
                return new SubmissionResponseDto
                {
                    Success = false,
                    SubmissionId = submissionId,
                    Message = genericErrorMsg,
                    Errors = new List<string> { "Internal server error occurred while writing to SharePoint." }
                };
            }
        }
    }
}
