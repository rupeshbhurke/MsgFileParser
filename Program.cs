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

            Console.WriteLine("Outlook MSG File Parser");
            Console.WriteLine("=======================\n");

            if (args.Length == 0)
            {
                Console.WriteLine("Usage: MsgFileParser <path_to_msg_file> [output_file_path]");
                Console.WriteLine("If output_file_path is not specified, a .txt file will be created in the same directory as the MSG file.");
                Console.WriteLine("You can specify either a complete file path or just a directory (ending with / or \\).");
                return;
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
                    Console.WriteLine("Error: MSG file path cannot be empty");
                    return;
                }

                if (!File.Exists(msgFilePath))
                {
                    Console.WriteLine($"Error: File not found - {msgFilePath}");
                    return;
                }

                // Validate file extension
                string fileExtension = Path.GetExtension(msgFilePath).ToLowerInvariant();
                if (fileExtension != ".msg")
                {
                    Console.WriteLine($"Warning: File does not have .msg extension - {fileExtension}");
                    Console.WriteLine("Attempting to process anyway...");
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
                    Console.WriteLine($"Error: Access denied to file - {msgFilePath}");
                    return;
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"Error: Cannot access file - {ex.Message}");
                    return;
                }

                // Check file size (warn for very large files)
                var fileInfo = new FileInfo(msgFilePath);
                if (fileInfo.Length > 50 * 1024 * 1024) // 50MB
                {
                    Console.WriteLine($"Warning: Large file detected ({fileInfo.Length / (1024 * 1024)}MB). Processing may take longer...");
                }

                // Validate and prepare output path
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    Console.WriteLine("Error: Output path cannot be empty");
                    return;
                }

                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Console.WriteLine($"Error: Output directory not found - {outputDir}");
                    return;
                }

                // Check if output file already exists and warn
                if (File.Exists(outputPath))
                {
                    Console.WriteLine($"Warning: Output file already exists and will be overwritten - {outputPath}");
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
                    Console.WriteLine($"Error: No write permission for output directory - {outputDir ?? "."}");
                    return;
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"Error: Cannot write to output directory - {ex.Message}");
                    return;
                }

                ProcessMsgFile(msgFilePath, outputPath);
                Console.WriteLine("Processing completed successfully.");
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine($"Error: Invalid file path - {ex.Message}");
            }
            catch (NotSupportedException ex)
            {
                Console.WriteLine($"Error: Operation not supported - {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                Console.WriteLine("Please check the input file and try again.");
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
                            Console.WriteLine("Using plain text body");
                        }
                        else if (!string.IsNullOrEmpty(msg.BodyHtml))
                        {
                            // If no plain text body, use HTML body
                            bodyText = msg.BodyHtml;
                            Console.WriteLine("Using HTML body (contains HTML formatting)");
                        }
                        else if (!string.IsNullOrEmpty(msg.BodyRtf))
                        {
                            // If no plain text or HTML, try RTF body
                            bodyText = msg.BodyRtf;
                            Console.WriteLine("Using RTF body (contains RTF formatting)");
                        }
                        else
                        {
                            bodyText = "[No body content found]";
                            Console.WriteLine("No body content found in any format");
                        }
                    }
                    catch (Exception ex)
                    {
                        bodyText = $"[Error reading email body: {ex.Message}]";
                        Console.WriteLine($"Warning: Could not read email body - {ex.Message}");
                    }
                    
                    content.AppendLine(bodyText);

                    // Write to file with error handling
                    try
                    {
                        File.WriteAllText(outputPath, content.ToString(), Encoding.UTF8);
                        Console.WriteLine($"Created: {outputPath}");
                    }
                    catch (IOException ex)
                    {
                        throw new IOException($"Failed to write output file: {ex.Message}", ex);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        throw new UnauthorizedAccessException($"Access denied writing to: {outputPath}", ex);
                    }
                }
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidDataException($"Invalid MSG file format: {ex.Message}", ex);
            }
            catch (OutOfMemoryException ex)
            {
                throw new OutOfMemoryException($"File too large to process: {msgFilePath}", ex);
            }
            catch (Exception ex) when (ex.Message.Contains("not a valid compound document"))
            {
                throw new InvalidDataException($"File is not a valid MSG file: {msgFilePath}", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to process MSG file: {ex.Message}", ex);
            }
        }
    }
}