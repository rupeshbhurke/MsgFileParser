namespace MsgFileParser
{
    public class MsgExportService
    {
        public void ExportToText(MsgFileInfo info, string outputPath)
        {
            var content = new System.Text.StringBuilder();
            content.AppendLine($"Subject: {info.Subject}");
            if (!string.IsNullOrEmpty(info.SenderDisplayName) && !string.IsNullOrEmpty(info.SenderEmail))
                content.AppendLine($"From: {info.SenderDisplayName} <{info.SenderEmail}>");
            else if (!string.IsNullOrEmpty(info.SenderEmail))
                content.AppendLine($"From: <{info.SenderEmail}>");
            else if (!string.IsNullOrEmpty(info.SenderDisplayName))
                content.AppendLine($"From: {info.SenderDisplayName}");
            else
                content.AppendLine("From: [Unknown]");
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

            // Embed images as Base64 in HTML
            if (info.ImageAttachments != null && info.ImageAttachments.Count > 0)
            {
                foreach (var img in info.ImageAttachments)
                {
                    string cidRef = $"cid:{img.ContentId}";
                    string base64Tag = $"data:{img.MimeType};base64,{img.Base64Data}";
                    // Replace all src="cid:..." with src="data:..."
                    htmlBody = System.Text.RegularExpressions.Regex.Replace(
                        htmlBody,
                        $"src=[\"']{cidRef}[\"']",
                        $"src=\"{base64Tag}\"",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    );
                }
            }

            // Build classic email header block
            string from = "";
            if (!string.IsNullOrEmpty(info.SenderDisplayName) && !string.IsNullOrEmpty(info.SenderEmail))
                from = $"{System.Net.WebUtility.HtmlEncode(info.SenderDisplayName)} &lt;{System.Net.WebUtility.HtmlEncode(info.SenderEmail)}&gt;";
            else if (!string.IsNullOrEmpty(info.SenderEmail))
                from = $"&lt;{System.Net.WebUtility.HtmlEncode(info.SenderEmail)}&gt;";
            else if (!string.IsNullOrEmpty(info.SenderDisplayName))
                from = System.Net.WebUtility.HtmlEncode(info.SenderDisplayName);
            string sent = System.Net.WebUtility.HtmlEncode(info.Date ?? "");
            string subject = System.Net.WebUtility.HtmlEncode(info.Subject ?? "MSG Email");

            // To, Cc fields (fallback: list all recipients in To, leave Cc blank)
            string to = info.Recipients != null && info.Recipients.Count > 0
                ? string.Join("; ", info.Recipients.ConvertAll(r => $"{System.Net.WebUtility.HtmlEncode(r.DisplayName)} &lt;{System.Net.WebUtility.HtmlEncode(r.Email)}&gt;"))
                : "";
            string cc = ""; // No Cc info available in MsgFileInfo

            string headerBlock = $@"
<div class='email-header' style='background:#f8f8f8; border:1px solid #ddd; padding:1em; margin-bottom:1em;'>
    <div><strong>From:</strong> {from}</div>
    <div><strong>Sent:</strong> {sent}</div>
    {(string.IsNullOrEmpty(to) ? "" : $"<div><strong>To:</strong> {to}</div>")}
    {(string.IsNullOrEmpty(cc) ? "" : $"<div><strong>Cc:</strong> {cc}</div>")}
    <div><strong>Subject:</strong> {subject}</div>
</div>
";

            string html = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>{subject}</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 2em; }}
        .email-header {{ background:#f8f8f8; border:1px solid #ddd; padding:1em; margin-bottom:1em; }}
        .body {{ border-top: 1px solid #ccc; padding-top: 1em; }}
    </style>
</head>
<body>
    {headerBlock}
    <div class='body'>
        {htmlBody}
    </div>
</body>
</html>";
            System.IO.File.WriteAllText(outputPath, html);
        }
    }
}
