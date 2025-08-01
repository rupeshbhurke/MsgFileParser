using System;
using System.IO;
using System.Text;
using MsgReader;
using MsgReader.Outlook;

namespace MsgFileParser
{
    class Program
    {
        static void Main(string[] args)
        {
            // Register code pages for proper encoding handling
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.WriteLine("MSG File Parser v1.0");
            Console.WriteLine("===================");

            if (args.Length == 0)
            {
                Console.WriteLine("ERROR: INVALID_USAGE - No input file specified");
                Console.WriteLine("USAGE: MsgFileParser <path_to_msg_file> [output_file_path]");
                Console.WriteLine("INFO: If output_file_path is not specified, a .txt file will be created in the same directory as the MSG file.");
                Console.WriteLine("INFO: You can specify either a complete file path or just a directory (ending with / or \\).");
                Environment.Exit(1);
            }

            string msgFilePath = args[0];
            string outputPath;
            if (args.Length > 1)
            {
                outputPath = args[1];
                if (Directory.Exists(args[1]) || args[1].EndsWith("/") || args[1].EndsWith("\\"))
                {
                    string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                    outputPath = Path.Combine(args[1].TrimEnd('/', '\\'), $"{fileName}.txt");
                }
            }
            else
            {
                string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                string directory = Path.GetDirectoryName(msgFilePath) ?? ".";
                outputPath = Path.Combine(directory, $"{fileName}.txt");
            }

            try
            {
                // ...existing validation code...
                if (string.IsNullOrWhiteSpace(msgFilePath))
                {
                    Console.WriteLine("ERROR: INVALID_INPUT - MSG file path cannot be empty");
                    Environment.Exit(2);
                }
                if (!File.Exists(msgFilePath))
                {
                    Console.WriteLine($"ERROR: FILE_NOT_FOUND - {msgFilePath}");
                    Environment.Exit(3);
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
                    Console.WriteLine($"ERROR: ACCESS_DENIED - Cannot read file: {msgFilePath}");
                    Environment.Exit(4);
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"ERROR: IO_ERROR - Cannot access file: {ex.Message}");
                    Environment.Exit(5);
                }
                var fileInfo = new FileInfo(msgFilePath);
                if (fileInfo.Length > 50 * 1024 * 1024)
                {
                    Console.WriteLine($"WARNING: LARGE_FILE - Size: {fileInfo.Length / (1024 * 1024)}MB, processing may take longer");
                }
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    Console.WriteLine("ERROR: INVALID_OUTPUT - Output path cannot be empty");
                    Environment.Exit(6);
                }
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Console.WriteLine($"ERROR: DIRECTORY_NOT_FOUND - Output directory does not exist: {outputDir}");
                    Environment.Exit(7);
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
                    Console.WriteLine($"ERROR: WRITE_ACCESS_DENIED - No write permission for directory: {outputDir ?? "."}");
                    Environment.Exit(8);
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"ERROR: WRITE_IO_ERROR - Cannot write to output directory: {ex.Message}");
                    Environment.Exit(9);
                }

                // Refactored: Use MsgFileParser and MsgExportService
                var parser = new MsgFileParser();
                var exportService = new MsgExportService();
                MsgFileInfo info = parser.Parse(msgFilePath);
                exportService.ExportToText(info, outputPath);
                Console.WriteLine($"SUCCESS: Processing completed - Output: {outputPath}");
                Environment.Exit(0);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"ERROR: INVALID_PATH - {ex.Message}");
                Environment.Exit(10);
            }
            catch (NotSupportedException ex)
            {
                Console.WriteLine($"ERROR: NOT_SUPPORTED - {ex.Message}");
                Environment.Exit(11);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: UNEXPECTED - {ex.Message}");
                Environment.Exit(99);
            }
        }

        static void ProcessMsgFile(string msgFilePath, string outputPath)
        {
            // ...removed, now handled by MsgFileParser and MsgExportService...
        }
    }
}