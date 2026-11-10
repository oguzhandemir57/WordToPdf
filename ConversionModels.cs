namespace WordToPdfApp;

public enum ConversionEngine
{
    None,
    MicrosoftWord,
    LibreOffice
}

public class ConversionProgress
{
    public int CurrentIndex { get; set; }
    public int TotalCount { get; set; }
    public string CurrentFileName { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public bool IsSuccess { get; set; }
    public string? ErrorMessage { get; set; }
}
