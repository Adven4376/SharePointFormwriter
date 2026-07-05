using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using RequestSubmissionFunctionApp.Configuration;
using RequestSubmissionFunctionApp.DTOs;
using RequestSubmissionFunctionApp.Interfaces;
using RequestSubmissionFunctionApp.Utilities;

namespace RequestSubmissionFunctionApp.Functions
{
    /// <summary>
    /// HTTP Trigger function that acts as the entry point for submissions from the Outlook Add-in.
    /// </summary>
    public class SubmitRequestFunction
    {
        private readonly ISubmissionService _submissionService;
        private readonly ApplicationSettings _appSettings;
        private readonly ILogger<SubmitRequestFunction> _logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubmitRequestFunction"/> class.
        /// </summary>
        public SubmitRequestFunction(
            ISubmissionService submissionService,
            ApplicationSettings appSettings,
            ILogger<SubmitRequestFunction> logger)
        {
            _submissionService = submissionService ?? throw new ArgumentNullException(nameof(submissionService));
            _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// HTTP Post endpoint to submit request forms and attachments.
        /// </summary>
        /// <param name="req">The HTTP request data.</param>
        /// <returns>The HTTP response containing the submission result.</returns>
        [Function("SubmitRequest")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "submit")] HttpRequestData req)
        {
            _logger.LogInformation("HTTP Trigger 'SubmitRequest' received a request.");

            // 1. Verify Content-Type Header
            if (!req.Headers.TryGetValues("Content-Type", out var contentTypeHeaders))
            {
                _logger.LogWarning("Request rejected: Content-Type header is missing.");
                var errorResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                await errorResponse.WriteAsJsonAsync(new { Error = "Content-Type header is required." });
                return errorResponse;
            }

            var contentType = contentTypeHeaders.FirstOrDefault();
            if (string.IsNullOrEmpty(contentType) || !contentType.StartsWith("multipart/form-data", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogWarning("Request rejected: Content-Type must be multipart/form-data. Received: {ContentType}", contentType);
                var errorResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                await errorResponse.WriteAsJsonAsync(new { Error = "Content-Type must be multipart/form-data." });
                return errorResponse;
            }

            try
            {
                // 2. Parse request multipart stream using the utility parser with configured size limits
                _logger.LogInformation("Parsing multipart request body stream...");
                SubmissionRequestDto requestDto = await MultipartParser.ParseAsync(req.Body, contentType, _appSettings.MaximumFileSize);

                // 3. Delegate execution to workflow SubmissionService
                SubmissionResponseDto responseDto = await _submissionService.ProcessSubmission(requestDto);

                // 4. Return appropriate HTTP response (aligning with Success property)
                var httpStatusCode = responseDto.Success ? HttpStatusCode.OK : HttpStatusCode.BadRequest;
                var response = req.CreateResponse(httpStatusCode);
                await response.WriteAsJsonAsync(responseDto);
                return response;
            }
            catch (ArgumentException argEx)
            {
                _logger.LogWarning(argEx, "Invalid arguments in incoming request.");
                var errorResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                await errorResponse.WriteAsJsonAsync(new { Error = argEx.Message });
                return errorResponse;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error occurred during HTTP SubmitRequest execution.");
                
                // Security: Do not leak internal exception details to the caller
                var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errorResponse.WriteAsJsonAsync(new { Error = "An unexpected error occurred while processing your request. Please try again." });
                return errorResponse;
            }
        }
    }
}
