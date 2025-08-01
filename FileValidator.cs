using System;
using System.IO;

namespace MsgFileParser
{
    public class FileValidator
    {
        public void ValidateInputFile(string msgFilePath)
        {
            if (string.IsNullOrWhiteSpace(msgFilePath))
            {
                throw new ArgumentException("MSG file path cannot be empty");
            }
            if (!File.Exists(msgFilePath))
            {
                throw new FileNotFoundException($"File not found: {msgFilePath}");
            }
            string fileExtension = Path.GetExtension(msgFilePath).ToLowerInvariant();
            if (fileExtension != ".msg")
            {
                Console.WriteLine($"WARNING: INVALID_EXTENSION - File has {fileExtension} extension, expected .msg");
                Console.WriteLine("INFO: Attempting to process anyway...");
            }
            try
            {
                using (var stream = File.OpenRead(msgFilePath)) { }
            }
            catch (UnauthorizedAccessException)
            {
                throw new UnauthorizedAccessException($"Cannot read file: {msgFilePath}");
            }
            catch (IOException ex)
            {
                throw new IOException($"Cannot access file: {ex.Message}");
            }
            var fileInfo = new FileInfo(msgFilePath);
            if (fileInfo.Length > 50 * 1024 * 1024)
            {
                Console.WriteLine($"WARNING: LARGE_FILE - Size: {fileInfo.Length / (1024 * 1024)}MB, processing may take longer");
            }
        }

        public void ValidateOutputPath(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("Output path cannot be empty");
            }
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                throw new DirectoryNotFoundException($"Output directory does not exist: {outputDir}");
            }
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"WARNING: FILE_EXISTS - Output file will be overwritten: {outputPath}");
            }
            try
            {
                string testFile = Path.Combine(outputDir ?? ".", "temp_write_test.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
            }
            catch (UnauthorizedAccessException)
            {
                throw new UnauthorizedAccessException($"No write permission for directory: {outputDir ?? "."}");
            }
            catch (IOException ex)
            {
                throw new IOException($"Cannot write to output directory: {ex.Message}");
            }
        }
    }
}
