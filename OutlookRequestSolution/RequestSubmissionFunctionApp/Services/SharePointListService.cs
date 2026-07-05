using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using RequestSubmissionFunctionApp.Configuration;
using RequestSubmissionFunctionApp.DTOs;
using RequestSubmissionFunctionApp.Interfaces;

namespace RequestSubmissionFunctionApp.Services
{
    /// <summary>
    /// Implementation of ISharePointListService using Microsoft Graph SDK v5.
    /// </summary>
    public class SharePointListService : ISharePointListService
    {
        private readonly GraphServiceClient _graphServiceClient;
        private readonly SharePointSettings _settings;
        private readonly ILogger<SharePointListService> _logger;

        public SharePointListService(
            GraphServiceClient graphServiceClient,
            SharePointSettings settings,
            ILogger<SharePointListService> logger)
        {
            _graphServiceClient = graphServiceClient ?? throw new ArgumentNullException(nameof(graphServiceClient));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<string> CreateSubmissionItem(Guid submissionId, SubmissionRequestDto request, string folderUrl)
        {
            _logger.LogInformation("Creating SharePoint list item for submission GUID {SubmissionId}.", submissionId);

            try
            {
                var listItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "Title", submissionId.ToString() }, // Standard Title field, maps to SubmissionId
                            { "SubmissionId", submissionId.ToString() },
                            { "Name", request.Name },
                            { "Email", request.Email },
                            { "Department", request.Department },
                            { "Description", request.Description },
                            { "FolderURL", folderUrl },
                            { "CreatedDate", DateTimeOffset.UtcNow.ToString("o") } // Store as ISO-8601 string
                        }
                    }
                };

                var createdItem = await _graphServiceClient.Sites[_settings.SiteId]
                    .Lists[_settings.ListId]
                    .Items
                    .PostAsync(listItem);

                if (createdItem == null || string.IsNullOrEmpty(createdItem.Id))
                {
                    throw new Exception("ListItem creation failed. No ID returned.");
                }

                _logger.LogInformation("Successfully created SharePoint List Item with ID: {Id}", createdItem.Id);
                return createdItem.Id;
            }
            catch (ODataError ex)
            {
                _logger.LogError(ex, "OData error creating SharePoint list item. Error: {Message}", ex.Error?.Message);
                throw new Exception("Error writing metadata to SharePoint List.", ex);
            }
        }

        /// <inheritdoc/>
        public async Task UpdateSubmission(string itemId, SubmissionRequestDto request, string folderUrl)
        {
            _logger.LogInformation("Updating SharePoint list item ID {ItemId}.", itemId);

            try
            {
                var fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "Name", request.Name },
                        { "Email", request.Email },
                        { "Department", request.Department },
                        { "Description", request.Description },
                        { "FolderURL", folderUrl }
                    }
                };

                await _graphServiceClient.Sites[_settings.SiteId]
                    .Lists[_settings.ListId]
                    .Items[itemId]
                    .Fields
                    .PatchAsync(fields);

                _logger.LogInformation("Successfully updated SharePoint List Item {ItemId}.", itemId);
            }
            catch (ODataError ex)
            {
                _logger.LogError(ex, "OData error updating list item. Error: {Message}", ex.Error?.Message);
                throw new Exception("Error updating metadata in SharePoint List.", ex);
            }
        }

        /// <inheritdoc/>
        public async Task DeleteSubmission(string itemId)
        {
            _logger.LogInformation("Deleting SharePoint list item ID {ItemId}.", itemId);

            try
            {
                await _graphServiceClient.Sites[_settings.SiteId]
                    .Lists[_settings.ListId]
                    .Items[itemId]
                    .DeleteAsync();

                _logger.LogInformation("Successfully deleted SharePoint List Item {ItemId}.", itemId);
            }
            catch (ODataError ex)
            {
                _logger.LogError(ex, "OData error deleting list item. Error: {Message}", ex.Error?.Message);
                throw new Exception("Error deleting metadata from SharePoint List.", ex);
            }
        }

        /// <inheritdoc/>
        public async Task<ListItem> GetSubmission(string itemId)
        {
            _logger.LogInformation("Retrieving SharePoint list item ID {ItemId}.", itemId);

            try
            {
                var item = await _graphServiceClient.Sites[_settings.SiteId]
                    .Lists[_settings.ListId]
                    .Items[itemId]
                    .GetAsync(config =>
                    {
                        config.QueryParameters.Expand = new[] { "fields" };
                    });

                if (item == null)
                {
                    throw new KeyNotFoundException($"ListItem with ID {itemId} not found.");
                }

                return item;
            }
            catch (ODataError ex)
            {
                _logger.LogError(ex, "OData error fetching list item. Error: {Message}", ex.Error?.Message);
                throw new Exception("Error reading metadata from SharePoint List.", ex);
            }
        }
    }
}
