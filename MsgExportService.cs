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

        public void ExportToHtml(MsgFileInfo info, string outputPath)
        {
            string htmlBody = info.BodyHtml ?? string.Empty;
            if (string.IsNullOrEmpty(htmlBody))
            {
                // Fallback: generate simple HTML from text
                var body = !string.IsNullOrEmpty(info.BodyText) ? info.BodyText : (!string.IsNullOrEmpty(info.BodyRtf) ? info.BodyRtf : "[No body content found]");
                htmlBody = $"<pre>{System.Net.WebUtility.HtmlEncode(body)}</pre>";
            }
            string recipientsHtml = info.Recipients.Count > 0
                ? string.Join("<br>", info.Recipients.ConvertAll(r => $"- {System.Net.WebUtility.HtmlEncode(r.Email)} ({System.Net.WebUtility.HtmlEncode(r.DisplayName)})"))
                : "[None found]";
            string html = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>{System.Net.WebUtility.HtmlEncode(info.Subject ?? "MSG Email")}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 2em; }}
        .meta {{ margin-bottom: 1em; }}
        .recipients {{ margin-bottom: 1em; }}
        .body {{ border-top: 1px solid #ccc; padding-top: 1em; }}
    </style>
</head>
<body>
    <div class='meta'>
        <strong>Subject:</strong> {System.Net.WebUtility.HtmlEncode(info.Subject)}<br>
        <strong>From:</strong> {System.Net.WebUtility.HtmlEncode(info.Sender)}<br>
        <strong>Date:</strong> {System.Net.WebUtility.HtmlEncode(info.Date)}<br>
    </div>
    <div class='recipients'>
        <strong>Recipients:</strong><br>
        {recipientsHtml}
    </div>
    <div class='body'>
        {htmlBody}
    </div>
</body>
</html>";
            System.IO.File.WriteAllText(outputPath, html);
        }
    }
}
