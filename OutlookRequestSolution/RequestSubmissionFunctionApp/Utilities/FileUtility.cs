using System;
using System.Collections.Generic;
using System.IO;

namespace RequestSubmissionFunctionApp.Utilities
{
    /// <summary>
    /// Utility class for performing common file operations and conversion helpers.
    /// </summary>
    public static class FileUtility
    {
        private static readonly Dictionary<string, string> MimeTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            { ".pdf", "application/pdf" },
            { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            { ".doc", "application/msword" },
            { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            { ".xls", "application/vnd.ms-excel" },
            { ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
            { ".ppt", "application/vnd.ms-powerpoint" },
            { ".png", "image/png" },
            { ".jpg", "image/jpeg" },
            { ".jpeg", "image/jpeg" },
            { ".txt", "text/plain" },
            { ".zip", "application/zip" }
        };

        /// <summary>
        /// Retrieves the file extension including the dot (e.g., ".pdf").
        /// </summary>
        public static string GetExtension(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return string.Empty;
            return Path.GetExtension(fileName).ToLowerInvariant();
        }

        /// <summary>
        /// Resolves the MIME type of a file based on its extension.
        /// </summary>
        public static string GetMimeType(string fileName)
        {
            var ext = GetExtension(fileName);
            if (MimeTypes.TryGetValue(ext, out var mimeType))
            {
                return mimeType;
            }
            return "application/octet-stream";
        }

        /// <summary>
        /// Generates a unique filename by appending a GUID.
        /// </summary>
        public static string GenerateUniqueFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return Guid.NewGuid().ToString();

            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            return $"{nameWithoutExt}_{Guid.NewGuid():N}{ext}";
        }

        /// <summary>
        /// Converts a byte count to a human-readable size (e.g., "350.2 KB" or "1.5 MB").
        /// </summary>
        public static string ConvertSize(long bytes)
        {
            string[] suffixes = { "Bytes", "KB", "MB", "GB", "TB" };
            int counter = 0;
            decimal number = bytes;
            
            while (Math.Round(number / 1024) >= 1)
            {
                number /= 1024;
                counter++;
            }
            
            return $"{number:n2} {suffixes[counter]}";
        }
    }
}
