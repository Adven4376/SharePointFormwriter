using System;
using System.Threading.Tasks;
using Microsoft.Graph.Models;
using RequestSubmissionFunctionApp.DTOs;

namespace RequestSubmissionFunctionApp.Interfaces
{
    /// <summary>
    /// Service responsible for managing metadata entries in the SharePoint List.
    /// </summary>
    public interface ISharePointListService
    {
        /// <summary>
        /// Creates a new SharePoint List item with submission details.
        /// </summary>
        /// <param name="submissionId">The unique ID of the submission.</param>
        /// <param name="request">The submission data.</param>
        /// <param name="folderUrl">The URL of the SharePoint folder containing attachments.</param>
        /// <returns>The ID of the newly created list item.</returns>
        Task<string> CreateSubmissionItem(Guid submissionId, SubmissionRequestDto request, string folderUrl);

        /// <summary>
        /// Updates an existing SharePoint List item.
        /// </summary>
        /// <param name="itemId">The ID of the list item to update.</param>
        /// <param name="request">The updated submission data.</param>
        /// <param name="folderUrl">The updated folder URL.</param>
        Task UpdateSubmission(string itemId, SubmissionRequestDto request, string folderUrl);

        /// <summary>
        /// Deletes a SharePoint List item.
        /// </summary>
        /// <param name="itemId">The ID of the list item to delete.</param>
        Task DeleteSubmission(string itemId);

        /// <summary>
        /// Retrieves a SharePoint List item by its ID.
        /// </summary>
        /// <param name="itemId">The ID of the list item to fetch.</param>
        /// <returns>The ListItem object returned by Microsoft Graph.</returns>
        Task<ListItem> GetSubmission(string itemId);
    }
}
