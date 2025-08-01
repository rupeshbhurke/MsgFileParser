namespace MsgFileParser
{
    public class MsgFileInfo
    {
        public string? Subject { get; set; }
        public string? Sender { get; set; }
        public string? Date { get; set; }
        public List<(string Email, string DisplayName)> Recipients { get; set; } = new();
        public string? BodyText { get; set; }
        public string? BodyHtml { get; set; }
        public string? BodyRtf { get; set; }
    }
}
