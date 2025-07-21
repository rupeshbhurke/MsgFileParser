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
                // User specified output path
                outputPath = args[1];
                
                // If they specified a directory, auto-generate filename
                if (Directory.Exists(args[1]) || args[1].EndsWith("/") || args[1].EndsWith("\\"))
                {
                    string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                    outputPath = Path.Combine(args[1].TrimEnd('/', '\\'), $"{fileName}.txt");
                }
            }
            else
            {
                // Auto-generate in same directory as MSG file
                string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                string directory = Path.GetDirectoryName(msgFilePath) ?? ".";
                outputPath = Path.Combine(directory, $"{fileName}.txt");
            }

            try
            {
                // Validate input file
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

                // Validate file extension
                string fileExtension = Path.GetExtension(msgFilePath).ToLowerInvariant();
                if (fileExtension != ".msg")
                {
                    Console.WriteLine($"WARNING: INVALID_EXTENSION - File has {fileExtension} extension, expected .msg");
                    Console.WriteLine("INFO: Attempting to process anyway...");
                }

                // Check file access permissions
                try
                {
                    using (var stream = File.OpenRead(msgFilePath))
                    {
                        // Just check if we can open the file
                    }
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

                // Check file size (warn for very large files)
                var fileInfo = new FileInfo(msgFilePath);
                if (fileInfo.Length > 50 * 1024 * 1024) // 50MB
                {
                    Console.WriteLine($"WARNING: LARGE_FILE - Size: {fileInfo.Length / (1024 * 1024)}MB, processing may take longer");
                }

                // Validate and prepare output path
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    Console.WriteLine("ERROR: INVALID_OUTPUT - Output path cannot be empty");
                    Environment.Exit(6);
                }

                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Console.WriteLine($"ERROR: DIRECTORY_NOT_FOUND - Output directory does not exist: {outputDir}");
                    Environment.Exit(7);
                }

                // Check if output file already exists and warn
                if (File.Exists(outputPath))
                {
                    Console.WriteLine($"WARNING: FILE_EXISTS - Output file will be overwritten: {outputPath}");
                }

                // Check write permissions for output location
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

                ProcessMsgFile(msgFilePath, outputPath);
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
            try
            {
                // Read the MSG file
                using (var msg = new Storage.Message(msgFilePath))
                {
                    // Prepare the content
                    var content = new StringBuilder();
                    
                    // Handle subject safely
                    string subject = string.IsNullOrEmpty(msg.Subject) ? "[No Subject]" : msg.Subject;
                    content.AppendLine($"Subject: {subject}");
                    
                    // Handle sender safely
                    string sender = msg.Sender?.ToString() ?? "[Unknown Sender]";
                    content.AppendLine($"From: {sender}");
                    
                    // Handle date safely
                    string date = msg.SentOn?.ToString("yyyy-MM-dd HH:mm:ss") ?? "[Unknown Date]";
                    content.AppendLine($"Date: {date}");

                    // Add recipients with error handling
                    if (msg.Recipients?.Count > 0)
                    {
                        content.AppendLine("\nRecipients:");
                        foreach (var recipient in msg.Recipients)
                        {
                            try
                            {
                                string email = string.IsNullOrEmpty(recipient.Email) ? "[No Email]" : recipient.Email;
                                string displayName = string.IsNullOrEmpty(recipient.DisplayName) ? "[No Name]" : recipient.DisplayName;
                                content.AppendLine($"- {email} ({displayName})");
                            }
                            catch (Exception ex)
                            {
                                content.AppendLine($"- [Error reading recipient: {ex.Message}]");
                            }
                        }
                    }
                    else
                    {
                        content.AppendLine("\nRecipients: [None found]");
                    }

                    // Add body with enhanced error handling
                    content.AppendLine("\nBody:");
                    
                    string bodyText = "";
                    try
                    {
                        // Try to get the body text - MSG files can have different body formats
                        if (!string.IsNullOrEmpty(msg.BodyText))
                        {
                            bodyText = msg.BodyText;
                            Console.WriteLine("INFO: Using plain text body");
                        }
                        else if (!string.IsNullOrEmpty(msg.BodyHtml))
                        {
                            // If no plain text body, use HTML body
                            bodyText = msg.BodyHtml;
                            Console.WriteLine("INFO: Using HTML body (contains formatting)");
                        }
                        else if (!string.IsNullOrEmpty(msg.BodyRtf))
                        {
                            // If no plain text or HTML, try RTF body
                            bodyText = msg.BodyRtf;
                            Console.WriteLine("INFO: Using RTF body (contains formatting)");
                        }
                        else
                        {
                            bodyText = "[No body content found]";
                            Console.WriteLine("WARNING: NO_BODY_CONTENT - No body found in any format");
                        }
                    }
                    catch (Exception ex)
                    {
                        bodyText = $"[Error reading email body: {ex.Message}]";
                        Console.WriteLine($"WARNING: BODY_READ_ERROR - {ex.Message}");
                    }
                    
                    content.AppendLine(bodyText);

                    // Write to file with error handling
                    try
                    {
                        File.WriteAllText(outputPath, content.ToString(), Encoding.UTF8);
                        Console.WriteLine($"INFO: File created successfully: {outputPath}");
                    }
                    catch (IOException ex)
                    {
                        throw new IOException($"WRITE_FAILED: {ex.Message}", ex);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        throw new UnauthorizedAccessException($"WRITE_ACCESS_DENIED: {outputPath}", ex);
                    }
                }
            }
            catch (InvalidDataException ex)
            {
                Console.WriteLine($"ERROR: INVALID_MSG_FORMAT - {ex.Message}");
                Environment.Exit(20);
            }
            catch (OutOfMemoryException)
            {
                Console.WriteLine($"ERROR: OUT_OF_MEMORY - File too large: {msgFilePath}");
                Environment.Exit(21);
            }
            catch (Exception ex) when (ex.Message.Contains("not a valid compound document") || 
                                      ex.Message.Contains("Invalid OLE structured storage file") ||
                                      ex.Message.Contains("OLE compound document"))
            {
                Console.WriteLine($"ERROR: INVALID_MSG_FILE - Not a valid MSG file: {msgFilePath}");
                Environment.Exit(22);
            }
            catch (IOException ex) when (ex.Message.StartsWith("WRITE_FAILED"))
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Environment.Exit(23);
            }
            catch (UnauthorizedAccessException ex) when (ex.Message.StartsWith("WRITE_ACCESS_DENIED"))
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Environment.Exit(24);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: PROCESSING_FAILED - {ex.Message}");
                Environment.Exit(25);
            }
        }
    }
}