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
                Console.WriteLine("Usage: MsgFileParser <path_to_msg_file> [output_directory]");
                Console.WriteLine("If output_directory is not specified, the text file will be created in the same directory as the MSG file.");
                return;
            }

            string msgFilePath = args[0];
            string outputDirectory = args.Length > 1 ? args[1] : Path.GetDirectoryName(msgFilePath);

            try
            {
                if (!File.Exists(msgFilePath))
                {
                    Console.WriteLine($"Error: File not found - {msgFilePath}");
                    return;
                }

                if (!Directory.Exists(outputDirectory))
                {
                    Console.WriteLine($"Error: Output directory not found - {outputDirectory}");
                    return;
                }

                ProcessMsgFile(msgFilePath, outputDirectory);
                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file: {ex.Message}");
            }
        }

        static void ProcessMsgFile(string msgFilePath, string outputDirectory)
        {
            // Read the MSG file
            using (var msg = new Storage.Message(msgFilePath))
            {
                string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                string outputPath = Path.Combine(outputDirectory, $"{fileName}.txt");

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