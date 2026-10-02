namespace OfficeTools.Common.Errors;

/// <summary>工具回傳給 agent 的錯誤碼。新增或修改時同步更新 plan/04-common-security-errors.md。</summary>
public static class ErrorCodes
{
    public const string FileNotFound = "FILE_NOT_FOUND";
    public const string FileTooLarge = "FILE_TOO_LARGE";
    public const string FileExists = "FILE_EXISTS";
    public const string FileLocked = "FILE_LOCKED";
    public const string PathNotAllowed = "PATH_NOT_ALLOWED";
    public const string UnsupportedFormat = "UNSUPPORTED_FORMAT";
    public const string PasswordProtected = "PASSWORD_PROTECTED";
    public const string CorruptFile = "CORRUPT_FILE";
    public const string SessionNotFound = "SESSION_NOT_FOUND";
    public const string SessionExpired = "SESSION_EXPIRED";
    public const string UnsavedChanges = "UNSAVED_CHANGES";
    public const string UnsafeToOverwrite = "UNSAFE_TO_OVERWRITE";
    public const string SheetNotFound = "SHEET_NOT_FOUND";
    public const string InvalidRange = "INVALID_RANGE";
}
