namespace MsgFileParser
{
    public class MsgFileInfo
    {
        public string? Subject { get; set; }
        public string? SenderEmail { get; set; }
        public string? SenderDisplayName { get; set; }
        public string? Date { get; set; }
        public List<(string Email, string DisplayName)> Recipients { get; set; } = new();
        public string? BodyText { get; set; }
        public string? BodyHtml { get; set; }
        public string? BodyRtf { get; set; }
        // List of image attachments: (ContentId, MimeType, Base64Data)
        public List<(string ContentId, string MimeType, string Base64Data)> ImageAttachments { get; set; } = new();
    }
}
