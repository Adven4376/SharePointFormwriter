using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.AspNetCore.Http
{
    /// <summary>
    /// Concrete implementation of IFormFile backed by a MemoryStream or other stream source.
    /// </summary>
    public class FormFile : IFormFile
    {
        private readonly Stream _stream;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormFile"/> class.
        /// </summary>
        /// <param name="stream">The source stream containing file data.</param>
        /// <param name="length">The size of the file in bytes.</param>
        /// <param name="name">The form field key name.</param>
        /// <param name="fileName">The uploaded file's name.</param>
        /// <param name="contentType">The MIME content-type header.</param>
        public FormFile(Stream stream, long length, string name, string fileName, string contentType)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            Length = length;
            Name = name;
            FileName = fileName;
            ContentType = contentType;
            ContentDisposition = $"form-data; name=\"{name}\"; filename=\"{fileName}\"";
        }

        /// <inheritdoc/>
        public string ContentType { get; }

        /// <inheritdoc/>
        public string ContentDisposition { get; }

        /// <inheritdoc/>
        public long Length { get; }

        /// <inheritdoc/>
        public string Name { get; }

        /// <inheritdoc/>
        public string FileName { get; }

        /// <inheritdoc/>
        public Stream OpenReadStream()
        {
            if (_stream.CanSeek)
            {
                _stream.Position = 0;
            }
            return _stream;
        }

        /// <inheritdoc/>
        public void CopyTo(Stream target)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            OpenReadStream().CopyTo(target);
        }

        /// <inheritdoc/>
        public Task CopyToAsync(Stream target, CancellationToken cancellationToken = default)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            return OpenReadStream().CopyToAsync(target, cancellationToken);
        }
    }
}
