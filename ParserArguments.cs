using System;
using System.IO;

namespace MsgFileParser
{
    public class ParserArguments
    {
        public string MsgFilePath { get; }
        public string OutputPath { get; }
        public string ExportFormat { get; }

        public ParserArguments(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentException("No input file specified");

            string? msgFilePath = null;
            string? outputPathArg = null;
            string exportFormat = "text";
            foreach (var arg in args)
            {
                var lower = arg.ToLower();
                if (lower == "--html" || lower == "html")
                    exportFormat = "html";
                else if (lower == "--text" || lower == "text")
                    exportFormat = "text";
                else if (msgFilePath == null && lower.EndsWith(".msg"))
                    msgFilePath = arg;
                else if (outputPathArg == null)
                    outputPathArg = arg;
            }
            if (msgFilePath == null)
                throw new ArgumentException("No MSG file specified");

            string outputPath;
            if (outputPathArg != null && (Directory.Exists(outputPathArg) || outputPathArg.EndsWith("/") || outputPathArg.EndsWith("\\")))
            {
                string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                string ext = exportFormat == "html" ? ".html" : ".txt";
                outputPath = Path.Combine(outputPathArg.TrimEnd('/', '\\'), $"{fileName}{ext}");
            }
            else if (outputPathArg != null)
            {
                outputPath = outputPathArg;
            }
            else
            {
                string fileName = Path.GetFileNameWithoutExtension(msgFilePath);
                string directory = Path.GetDirectoryName(msgFilePath) ?? ".";
                string ext = exportFormat == "html" ? ".html" : ".txt";
                outputPath = Path.Combine(directory, $"{fileName}{ext}");
            }

            MsgFilePath = msgFilePath;
            OutputPath = outputPath;
            ExportFormat = exportFormat;
        }
    }
}
