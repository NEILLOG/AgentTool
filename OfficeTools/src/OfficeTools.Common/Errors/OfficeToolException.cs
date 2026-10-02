namespace OfficeTools.Common.Errors;

/// <summary>可預期的工具錯誤。<see cref="Hint"/> 要告訴 agent 下一步該怎麼做。</summary>
public sealed class OfficeToolException : Exception
{
    public OfficeToolException(string code, string message, string? hint = null, Exception? innerException = null)
        : base(message, innerException)
    {
        Code = code;
        Hint = hint;
    }

    public string Code { get; }

    public string? Hint { get; }
}
