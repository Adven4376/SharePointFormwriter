using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using RequestSubmissionFunctionApp.DTOs;

namespace RequestSubmissionFunctionApp.Interfaces
{
    /// <summary>
    /// Service responsible for managing attachments in the SharePoint Document Library.
    /// </summary>
    public interface ISharePointDocumentService
    {
        /// <summary>
        /// Creates a unique folder for the submission based on the GUID inside the Document Library.
        /// </summary>
        /// <param name="submissionId">The unique ID for the submission.</param>
        /// <returns>The Web URL of the created folder in SharePoint.</returns>
        Task<string> CreateSubmissionFolder(Guid submissionId);

        /// <summary>
        /// Uploads an individual attachment file to the submission folder.
        /// </summary>
        /// <param name="submissionId">The unique ID of the submission.</param>
        /// <param name="file">The file to upload.</param>
        /// <returns>The result of the file upload.</returns>
        Task<FileUploadResultDto> UploadFile(Guid submissionId, IFormFile file);

        /// <summary>
        /// Uploads multiple attachments into the submission folder in SharePoint.
        /// </summary>
        /// <param name="submissionId">The unique ID for the submission.</param>
        /// <param name="files">The collection of files to upload.</param>
        /// <returns>A list of results for each file indicating success or failure details.</returns>
        Task<List<FileUploadResultDto>> UploadMultipleFiles(Guid submissionId, List<IFormFile> files);

        /// <summary>
        /// Retrieves the absolute web URL of the folder for a given submission.
        /// </summary>
        /// <param name="submissionId">The unique ID for the submission.</param>
        /// <returns>The Web URL of the folder.</returns>
        Task<string> GetFolderUrl(Guid submissionId);

        /// <summary>
        /// Deletes specified uploaded files from the submission folder.
        /// </summary>
        /// <param name="submissionId">The unique ID for the submission.</param>
        /// <param name="fileNames">The names of the files to delete.</param>
        Task DeleteUploadedFiles(Guid submissionId, List<string> fileNames);

        /// <summary>
        /// Deletes the submission folder and all its contents from SharePoint (used for rollback).
        /// </summary>
        /// <param name="submissionId">The unique ID for the submission.</param>
        Task DeleteFolder(Guid submissionId);
    }
}
