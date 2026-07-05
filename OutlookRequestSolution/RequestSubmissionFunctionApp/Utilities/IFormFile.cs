using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.AspNetCore.Http
{
    /// <summary>
    /// Represents a file sent with the HttpRequest. Custom interface designed to avoid full framework references.
    /// </summary>
    public interface IFormFile
    {
        /// <summary>
        /// Gets the raw Content-Type header of the uploaded file.
        /// </summary>
        string ContentType { get; }

        /// <summary>
        /// Gets the raw Content-Disposition header of the uploaded file.
        /// </summary>
        string ContentDisposition { get; }

        /// <summary>
        /// Gets the file length in bytes.
        /// </summary>
        long Length { get; }

        /// <summary>
        /// Gets the form field name from the Content-Disposition header.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets the file name from the Content-Disposition header.
        /// </summary>
        string FileName { get; }

        /// <summary>
        /// Opens the request stream for reading the uploaded file.
        /// </summary>
        Stream OpenReadStream();

        /// <summary>
        /// Copies the contents of the uploaded file to the target stream.
        /// </summary>
        void CopyTo(Stream target);

        /// <summary>
        /// Asynchronously copies the contents of the uploaded file to the target stream.
        /// </summary>
        Task CopyToAsync(Stream target, CancellationToken cancellationToken = default);
    }
}
