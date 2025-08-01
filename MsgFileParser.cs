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
            return info;
        }
    }
}
