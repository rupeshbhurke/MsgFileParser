using MsgReader.Outlook;

namespace MsgFileParser
{
    public class MsgFileParser
    {
        public MsgFileInfo Parse(string msgFilePath)
        {
            using var msg = new Storage.Message(msgFilePath);
            var info = new MsgFileInfo
            {
                Subject = msg.Subject ?? "",
                SenderEmail = msg.Sender?.Email ?? "",
                SenderDisplayName = msg.Sender?.DisplayName ?? "",
                Date = msg.SentOn?.ToString("yyyy-MM-dd HH:mm:ss") ?? "",
                BodyText = msg.BodyText ?? "",
                BodyHtml = msg.BodyHtml ?? "",
                BodyRtf = msg.BodyRtf ?? ""
            };
            if (msg.Recipients != null)
            {
                foreach (var r in msg.Recipients)
                {
                    info.Recipients.Add((r.Email ?? "", r.DisplayName ?? ""));
                }
            }
            // Extract image attachments for embedding
            if (msg.Attachments != null)
            {
                foreach (var attObj in msg.Attachments)
                {
                    if (attObj is MsgReader.Outlook.Storage.Attachment att)
                    {
                        // Try to get mime type from FileExtension if MimeType is not available
                        string mimeType = "";
                        if (!string.IsNullOrEmpty(att.FileName))
                        {
                            var ext = System.IO.Path.GetExtension(att.FileName).ToLowerInvariant();
                            switch (ext)
                            {
                                case ".png": mimeType = "image/png"; break;
                                case ".jpg":
                                case ".jpeg": mimeType = "image/jpeg"; break;
                                case ".gif": mimeType = "image/gif"; break;
                                case ".bmp": mimeType = "image/bmp"; break;
                                default: mimeType = ""; break;
                            }
                        }
                        // Only embed images
                        if (!string.IsNullOrEmpty(mimeType))
                        {
                            var contentId = att.ContentId ?? att.FileName ?? "";
                            var data = att.Data;
                            if (data != null && data.Length > 0)
                            {
                                string base64 = Convert.ToBase64String(data);
                                info.ImageAttachments.Add((contentId, mimeType, base64));
                            }
                        }
                    }
                }
            }
            return info;
        }
    }
}
