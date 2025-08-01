namespace MsgFileParser
{
    /// <summary>
    /// Defines exit codes for the MsgFileParser application.
    /// </summary>
    public enum ExitCode
    {
        Success = 0,
        InvalidUsage = 1,
        FileNotFound = 3,
        AccessDenied = 4,
        IoError = 5,
        DirectoryNotFound = 7,
        InvalidPath = 10,
        NotSupported = 11,
        Unexpected = 99
    }
}
