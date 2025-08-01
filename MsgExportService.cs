namespace MsgFileParser
{
    public class MsgExportService
    {
        public void ExportToText(MsgFileInfo info, string outputPath)
        {
            var content = new System.Text.StringBuilder();
            content.AppendLine($"Subject: {info.Subject}");
            content.AppendLine($"From: {info.Sender}");
            content.AppendLine($"Date: {info.Date}");
            if (info.Recipients.Count > 0)
            {
                content.AppendLine("\nRecipients:");
                foreach (var r in info.Recipients)
                    content.AppendLine($"- {r.Email} ({r.DisplayName})");
            }
            else
            {
                content.AppendLine("\nRecipients: [None found]");
            }
            content.AppendLine("\nBody:");
            if (!string.IsNullOrEmpty(info.BodyText))
                content.AppendLine(info.BodyText);
            else if (!string.IsNullOrEmpty(info.BodyHtml))
                content.AppendLine(info.BodyHtml);
            else if (!string.IsNullOrEmpty(info.BodyRtf))
                content.AppendLine(info.BodyRtf);
            else
                content.AppendLine("[No body content found]");
            System.IO.File.WriteAllText(outputPath, content.ToString());
        }
    }
}
