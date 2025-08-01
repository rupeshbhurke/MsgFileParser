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
                Console.WriteLine("USAGE: MsgFileParser <path_to_msg_file> [output_file_path] [--text|--html]");
                Console.WriteLine("INFO: If output_file_path is not specified, a .txt file will be created in the same directory as the MSG file.");
                Console.WriteLine("INFO: You can specify either a complete file path or just a directory (ending with / or \"");
                Console.WriteLine("INFO: You can provide --text or --html in any position.");
                Environment.Exit((int)ExitCode.InvalidUsage);
            }

            ParserArguments parsedArgs;
            try
            {
                parsedArgs = new ParserArguments(args);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"ERROR: INVALID_USAGE - {ex.Message}");
                Console.WriteLine("USAGE: MsgFileParser <path_to_msg_file> [output_file_path] [--text|--html]");
                Environment.Exit((int)ExitCode.InvalidUsage);
                return;
            }

            string msgFilePath = parsedArgs.MsgFilePath;
            string outputPath = parsedArgs.OutputPath;
            string exportFormat = parsedArgs.ExportFormat;

            var validator = new FileValidator();
            try
            {
                validator.ValidateInputFile(msgFilePath);
                validator.ValidateOutputPath(outputPath);

                var parser = new MsgFileParser();
                var exportService = new MsgExportService();
                MsgFileInfo info = parser.Parse(msgFilePath);
                if (exportFormat == "html")
                {
                    exportService.ExportToHtml(info, outputPath);
                }
                else
                {
                    exportService.ExportToText(info, outputPath);
                }
                Console.WriteLine($"SUCCESS: Processing completed - Output: {outputPath}");
                Environment.Exit((int)ExitCode.Success);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"ERROR: INVALID_PATH - {ex.Message}");
                Environment.Exit((int)ExitCode.InvalidPath);
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"ERROR: FILE_NOT_FOUND - {ex.Message}");
                Environment.Exit((int)ExitCode.FileNotFound);
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine($"ERROR: DIRECTORY_NOT_FOUND - {ex.Message}");
                Environment.Exit((int)ExitCode.DirectoryNotFound);
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"ERROR: ACCESS_DENIED - {ex.Message}");
                Environment.Exit((int)ExitCode.AccessDenied);
            }
            catch (IOException ex)
            {
                Console.WriteLine($"ERROR: IO_ERROR - {ex.Message}");
                Environment.Exit((int)ExitCode.IoError);
            }
            catch (NotSupportedException ex)
            {
                Console.WriteLine($"ERROR: NOT_SUPPORTED - {ex.Message}");
                Environment.Exit((int)ExitCode.NotSupported);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: UNEXPECTED - {ex.Message}");
                Environment.Exit((int)ExitCode.Unexpected);
            }
        }

        // Removed legacy ProcessMsgFile method; all logic is now refactored.
    }
}