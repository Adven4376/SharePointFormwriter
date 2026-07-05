using System.Threading.Tasks;
using RequestSubmissionFunctionApp.DTOs;

namespace RequestSubmissionFunctionApp.Interfaces
{
    /// <summary>
    /// Core service representing the submission workflow orchestration.
    /// </summary>
    public interface ISubmissionService
    {
        /// <summary>
        /// Executes the submission workflow: Validates the request, uploads files to SharePoint, creates the list metadata item, and manages rollbacks.
        /// </summary>
        /// <param name="request">The submission request payload.</param>
        /// <returns>A submission response DTO containing execution metadata, tracking GUID, and results.</returns>
        Task<SubmissionResponseDto> ProcessSubmission(SubmissionRequestDto request);
    }
}
