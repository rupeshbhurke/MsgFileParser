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
                if (!File.Exists(msgFilePath))
                {
                    Console.WriteLine($"Error: File not found - {msgFilePath}");
                    return;
                }

                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Console.WriteLine($"Error: Output directory not found - {outputDir}");
                    return;
                }

                ProcessMsgFile(msgFilePath, outputPath);
                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file: {ex.Message}");
            }
        }

        static void ProcessMsgFile(string msgFilePath, string outputPath)
        {
            // Read the MSG file
            using (var msg = new Storage.Message(msgFilePath))
            {
                // Prepare the content
                var content = new StringBuilder();
                content.AppendLine($"Subject: {msg.Subject}");
                content.AppendLine($"From: {msg.Sender}");
                content.AppendLine($"Date: {msg.SentOn?.ToString("yyyy-MM-dd HH:mm:ss")}");

                // Add recipients
                if (msg.Recipients.Count > 0)
                {
                    content.AppendLine("\nRecipients:");
                    foreach (var recipient in msg.Recipients)
                    {
                        content.AppendLine($"- {recipient.Email} ({recipient.DisplayName})");
                    }
                }

                // Add body
                content.AppendLine("\nBody:");
                
                // Try to get the body text - MSG files can have different body formats
                string bodyText = "";
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
                
                content.AppendLine(bodyText);

                // Write to file
                File.WriteAllText(outputPath, content.ToString());
                Console.WriteLine($"Created: {outputPath}");
            }
        }
    }
}