using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;
using RequestSubmissionFunctionApp.DTOs;

namespace RequestSubmissionFunctionApp.Utilities
{
    /// <summary>
    /// Utility class to parse multipart/form-data content streams manually in a .NET Isolated Worker function.
    /// </summary>
    public static class MultipartParser
    {
        /// <summary>
        /// Parses a multipart request stream into a <see cref="SubmissionRequestDto"/>.
        /// </summary>
        /// <param name="body">The HTTP body stream.</param>
        /// <param name="contentType">The raw Content-Type header containing the boundary.</param>
        /// <returns>A populated SubmissionRequestDto.</returns>
        /// <exception cref="ArgumentException">Thrown when the content type or boundary is invalid.</exception>
        public static async Task<SubmissionRequestDto> ParseAsync(Stream body, string contentType)
        {
            var requestDto = new SubmissionRequestDto();

            // Parse media type header and extract the boundary
            if (!MediaTypeHeaderValue.TryParse(contentType, out var mediaTypeHeader))
            {
                throw new ArgumentException("Invalid Content-Type header.");
            }

            var boundary = HeaderUtilities.RemoveQuotes(mediaTypeHeader.Boundary).Value;
            if (string.IsNullOrEmpty(boundary))
            {
                throw new ArgumentException("Invalid multipart/form-data: missing boundary.");
            }

            var reader = new MultipartReader(boundary, body);
            MultipartSection? section;

            while ((section = await reader.ReadNextSectionAsync()) != null)
            {
                var hasContentDispositionHeader = ContentDispositionHeaderValue.TryParse(
                    section.ContentDisposition, out var contentDisposition);

                if (hasContentDispositionHeader && contentDisposition != null)
                {
                    var name = HeaderUtilities.RemoveQuotes(contentDisposition.Name).Value?.ToString() ?? string.Empty;

                    // Check if this part represents a file attachment
                    if (contentDisposition.DispositionType.Equals("form-data") && 
                        !string.IsNullOrEmpty(contentDisposition.FileName.Value))
                    {
                        var fileName = HeaderUtilities.RemoveQuotes(contentDisposition.FileName).Value?.ToString() ?? string.Empty;
                        
                        // Copy to MemoryStream since the section stream is transient
                        var memoryStream = new MemoryStream();
                        await section.Body.CopyToAsync(memoryStream);
                        memoryStream.Position = 0;

                        // Create our custom FormFile instance
                        var formFile = new FormFile(
                            memoryStream,
                            memoryStream.Length,
                            name,
                            fileName,
                            section.ContentType ?? "application/octet-stream"
                        );

                        requestDto.Attachments.Add(formFile);
                    }
                    else
                    {
                        // It is a standard form field text value
                        using var readerStream = new StreamReader(section.Body);
                        var value = await readerStream.ReadToEndAsync();

                        switch (name.ToLowerInvariant())
                        {
                            case "name":
                                requestDto.Name = value;
                                break;
                            case "email":
                                requestDto.Email = value;
                                break;
                            case "department":
                                requestDto.Department = value;
                                break;
                            case "description":
                                requestDto.Description = value;
                                break;
                        }
                    }
                }
            }

            return requestDto;
        }
    }
}
