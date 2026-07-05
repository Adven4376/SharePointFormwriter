using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using RequestSubmissionFunctionApp.Configuration;
using RequestSubmissionFunctionApp.DTOs;
using RequestSubmissionFunctionApp.Interfaces;

namespace RequestSubmissionFunctionApp.Services
{
    /// <summary>
    /// Implementation of ISharePointDocumentService utilizing Microsoft Graph SDK v5.
    /// </summary>
    public class SharePointDocumentService : ISharePointDocumentService
    {
        private readonly GraphServiceClient _graphServiceClient;
        private readonly SharePointSettings _settings;
        private readonly ILogger<SharePointDocumentService> _logger;

        public SharePointDocumentService(
            GraphServiceClient graphServiceClient,
            SharePointSettings settings,
            ILogger<SharePointDocumentService> logger)
        {
            _graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<string> CreateSubmissionFolder(Guid submissionId)
        {
            _logger.LogInformation("Creating submission folder for GUID {SubmissionId} in SharePoint.", submissionId);
            
            try
            {
                var drive = await GetDriveAsync();
                if (string.IsNullOrEmpty(drive.Id))
                {
                    throw new Exception("Resolved Drive ID is null or empty.");
                }

                await EnsureParentFolderExistsAsync(drive.Id);

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

                // Create the folder inside RequestAttachments
                var createdFolder = await _graphServiceClient.Drives[drive.Id]
                    .Root
                    .ItemWithPath(_settings.RequestAttachmentsFolder)
                    .Children
                    .PostAsync(newFolder);

                if (createdFolder == null || string.IsNullOrEmpty(createdFolder.WebUrl))
                {
                    throw new Exception($"Failed to retrieve WebUrl for created folder {submissionFolderName}.");
                }

                _logger.LogInformation("Successfully created SharePoint folder: {WebUrl}", createdFolder.WebUrl);
                return createdFolder.WebUrl;
            }
            catch (ODataError ex)
            {
                _logger.LogError(ex, "OData error creating folder in SharePoint. Error: {Message}", ex.Error?.Message);
                throw new Exception("Error interacting with SharePoint Document Library.", ex);
            }
        }

        /// <inheritdoc/>
        public async Task<FileUploadResultDto> UploadFile(Guid submissionId, IFormFile file)
        {
            var result = new FileUploadResultDto
            {
                FileName = file.FileName,
                FileSize = file.Length
            };

            try
            {
                var drive = await GetDriveAsync();
                var folderPath = $"{_settings.RequestAttachmentsFolder}/{submissionId}";
                var folderItem = await _graphServiceClient.Drives[drive.Id]
                    .Root
                    .ItemWithPath(folderPath)
                    .GetAsync();

                if (folderItem == null || string.IsNullOrEmpty(folderItem.Id))
                {
                    throw new FileNotFoundException($"Submission folder '{folderPath}' not found in SharePoint.");
                }

                _logger.LogInformation("Uploading file '{FileName}' ({Length} bytes) to SharePoint.", file.FileName, file.Length);

                using var stream = file.OpenReadStream();

                // Chunked upload session for Microsoft Graph
                var uploadSessionRequestBody = new CreateUploadSessionPostRequestBody
                {
                    Item = new DriveItemUploadableProperties
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "@microsoft.graph.conflictBehavior", "replace" }
                        }
                    }
                };

                var uploadSession = await _graphServiceClient.Drives[drive.Id]
                    .Items[folderItem.Id]
                    .ItemWithPath(file.FileName)
                    .CreateUploadSession
                    .PostAsync(uploadSessionRequestBody);

                if (uploadSession == null || string.IsNullOrEmpty(uploadSession.UploadUrl))
                {
                    throw new Exception("Failed to create upload session.");
                }

                int maxSliceSize = 320 * 1024; // 320 KB chunks
                var fileUploadTask = new LargeFileUploadTask<DriveItem>(
                    uploadSession,
                    stream,
                    maxSliceSize,
                    _graphServiceClient.RequestAdapter
                );

                var uploadResult = await fileUploadTask.UploadAsync();
                if (uploadResult.UploadSucceeded)
                {
                    result.UploadStatus = true;
                    result.FileUrl = uploadResult.ItemResponse.WebUrl ?? string.Empty;
                    _logger.LogInformation("Successfully uploaded file '{FileName}'. URL: {Url}", file.FileName, result.FileUrl);
                }
                else
                {
                    result.UploadStatus = false;
                    result.ErrorMessage = "Upload session did not complete successfully.";
                }
            }
            catch (ODataError ex)
            {
                result.UploadStatus = false;
                result.ErrorMessage = ex.Error?.Message ?? "SharePoint communication error.";
                _logger.LogError(ex, "OData error uploading file '{FileName}'.", file.FileName);
            }
            catch (Exception ex)
            {
                result.UploadStatus = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Error uploading file '{FileName}'.", file.FileName);
            }

            return result;
        }

        /// <inheritdoc/>
        public async Task<List<FileUploadResultDto>> UploadMultipleFiles(Guid submissionId, List<IFormFile> files)
        {
            _logger.LogInformation("Uploading {Count} files for submission {SubmissionId}.", files.Count, submissionId);
            var results = new List<FileUploadResultDto>();

            foreach (var file in files)
            {
                var result = await UploadFile(submissionId, file);
                results.Add(result);
            }

            return results;
        }

        /// <inheritdoc/>
        public async Task<string> GetFolderUrl(Guid submissionId)
        {
            try
            {
                var drive = await GetDriveAsync();
                var folderPath = $"{_settings.RequestAttachmentsFolder}/{submissionId}";
                var folderItem = await _graphServiceClient.Drives[drive.Id]
                    .Root
                    .ItemWithPath(folderPath)
                    .GetAsync();

                return folderItem?.WebUrl ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving WebUrl for submission folder {SubmissionId}.", submissionId);
                return string.Empty;
            }
        }

        /// <inheritdoc/>
        public async Task DeleteUploadedFiles(Guid submissionId, List<string> fileNames)
        {
            try
            {
                var drive = await GetDriveAsync();
                foreach (var fileName in fileNames)
                {
                    var filePath = $"{_settings.RequestAttachmentsFolder}/{submissionId}/{fileName}";
                    try
                    {
                        var item = await _graphServiceClient.Drives[drive.Id]
                            .Root
                            .ItemWithPath(filePath)
                            .GetAsync();

                        if (item != null && !string.IsNullOrEmpty(item.Id))
                        {
                            await _graphServiceClient.Drives[drive.Id]
                                .Items[item.Id]
                                .DeleteAsync();
                            _logger.LogInformation("Deleted file: {FilePath}", filePath);
                        }
                    }
                    catch (ODataError ex) when (ex.ResponseStatusCode == 404)
                    {
                        // File already deleted or missing
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting files for submission folder {SubmissionId}.", submissionId);
            }
        }

        /// <inheritdoc/>
        public async Task DeleteFolder(Guid submissionId)
        {
            _logger.LogWarning("Deleting attachment folder for submission {SubmissionId} due to workflow rollback.", submissionId);
            
            try
            {
                var drive = await GetDriveAsync();
                var folderPath = $"{_settings.RequestAttachmentsFolder}/{submissionId}";
                
                var folderItem = await _graphServiceClient.Drives[drive.Id]
                    .Root
                    .ItemWithPath(folderPath)
                    .GetAsync();

                if (folderItem != null && !string.IsNullOrEmpty(folderItem.Id))
                {
                    await _graphServiceClient.Drives[drive.Id]
                        .Items[folderItem.Id]
                        .DeleteAsync();
                    
                    _logger.LogInformation("Successfully deleted folder: {FolderPath}", folderPath);
                }
            }
            catch (ODataError ex) when (ex.ResponseStatusCode == 404)
            {
                _logger.LogInformation("Folder for deletion not found.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to delete folder {SubmissionId} from SharePoint.", submissionId);
            }
        }

        #region Helper Methods

        private async Task<Drive> GetDriveAsync()
        {
            if (string.IsNullOrWhiteSpace(_settings.SiteId))
            {
                throw new ArgumentException("SharePoint SiteId is not configured.");
            }

            if (string.IsNullOrWhiteSpace(_settings.LibraryId))
            {
                throw new ArgumentException("SharePoint LibraryId is not configured.");
            }

            var drivesCollection = await _graphServiceClient.Sites[_settings.SiteId].Drives.GetAsync();
            if (drivesCollection?.Value == null)
            {
                throw new Exception("No document libraries found on the SharePoint site.");
            }

            var drive = drivesCollection.Value.FirstOrDefault(d =>
                d.Id == _settings.LibraryId ||
                string.Equals(d.Name, _settings.LibraryId, StringComparison.OrdinalIgnoreCase));

            if (drive == null)
            {
                throw new DirectoryNotFoundException($"SharePoint Document Library '{_settings.LibraryId}' not found.");
            }

            return drive;
        }

        private async Task EnsureParentFolderExistsAsync(string driveId)
        {
            try
            {
                await _graphServiceClient.Drives[driveId]
                    .Root
                    .ItemWithPath(_settings.RequestAttachmentsFolder)
                    .GetAsync();
            }
            catch (ODataError ex) when (ex.ResponseStatusCode == 404)
            {
                _logger.LogInformation("Root folder '{FolderName}' not found. Creating it at the drive root.", _settings.RequestAttachmentsFolder);
                
                var newFolder = new DriveItem
                {
                    Name = _settings.RequestAttachmentsFolder,
                    Folder = new Folder()
                };

                await _graphServiceClient.Drives[driveId]
                    .Items["root"]
                    .Children
                    .PostAsync(newFolder);
            }
        }

        #endregion
    }
}
