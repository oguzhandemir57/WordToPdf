using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WordToPdfApp;

public class WordToPdfConverterService
{
    public static (ConversionEngine Engine, string Description, string? Path) DetectEngine()
    {
        // 1. Check Microsoft Word COM
        try
        {
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType != null)
            {
                return (ConversionEngine.MicrosoftWord, "Microsoft Word (Office COM)", null);
            }
        }
        catch
        {
            // Ignore error during check
        }

        // 2. Check LibreOffice
        string? libreOfficePath = FindLibreOffice();
        if (!string.IsNullOrEmpty(libreOfficePath) && File.Exists(libreOfficePath))
        {
            return (ConversionEngine.LibreOffice, "LibreOffice (soffice)", libreOfficePath);
        }

        return (ConversionEngine.None, "Dönüştürme motoru (Microsoft Word veya LibreOffice) bulunamadı.", null);
    }

    public static List<FileInfo> ScanWordFiles(string folderPath, bool includeSubfolders)
    {
        if (!Directory.Exists(folderPath))
        {
            return new List<FileInfo>();
        }

        var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var dirInfo = new DirectoryInfo(folderPath);

        // Desteklenen uzantılar: .docx, .doc, .docm
        // Word geçici/kilit dosyalarını (örn: ~$rapor.docx) hariç tut
        var files = dirInfo.EnumerateFiles("*.*", searchOption)
            .Where(f => !f.Name.StartsWith("~$", StringComparison.OrdinalIgnoreCase) &&
                       (f.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
                        f.Extension.Equals(".doc", StringComparison.OrdinalIgnoreCase) ||
                        f.Extension.Equals(".docm", StringComparison.OrdinalIgnoreCase)))
            .OrderBy(f => f.Name)
            .ToList();

        return files;
    }

    public static string GetDesktopPdfFolderPath()
    {
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        return Path.Combine(desktopPath, "PdfDosyaları");
    }

    public async Task ConvertFilesAsync(
        List<FileInfo> files,
        string targetFolder,
        IProgress<ConversionProgress> progress,
        CancellationToken cancellationToken)
    {
        if (!Directory.Exists(targetFolder))
        {
            Directory.CreateDirectory(targetFolder);
        }

        var (engine, desc, loPath) = DetectEngine();

        if (engine == ConversionEngine.MicrosoftWord)
        {
            await Task.Run(() => ConvertWithWordCom(files, targetFolder, progress, cancellationToken), cancellationToken);
        }
        else if (engine == ConversionEngine.LibreOffice && !string.IsNullOrEmpty(loPath))
        {
            await ConvertWithLibreOfficeAsync(files, targetFolder, loPath, progress, cancellationToken);
        }
        else
        {
            throw new InvalidOperationException(
                "Word dosyalarını PDF formatına dönüştürebilmek için sisteminizde Microsoft Word veya LibreOffice kurulu olmalıdır.\n\n" +
                "Lütfen bilgisayarınızda Microsoft Office/Word kurulu olduğundan emin olun.");
        }
    }

    private void ConvertWithWordCom(
        List<FileInfo> files,
        string targetFolder,
        IProgress<ConversionProgress> progress,
        CancellationToken cancellationToken)
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application")
            ?? throw new InvalidOperationException("Microsoft Word COM bileşeni bulunamadı.");

        dynamic? wordApp = Activator.CreateInstance(wordType)
            ?? throw new InvalidOperationException("Microsoft Word başlatılamadı.");

        try
        {
            wordApp.Visible = false;
            wordApp.DisplayAlerts = 0; // wdAlertsNone

            for (int i = 0; i < files.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var file = files[i];
                string outputPdfPath = Path.Combine(targetFolder, Path.GetFileNameWithoutExtension(file.Name) + ".pdf");

                progress.Report(new ConversionProgress
                {
                    CurrentIndex = i + 1,
                    TotalCount = files.Count,
                    CurrentFileName = file.Name,
                    Status = "Dönüştürülüyor...",
                    IsSuccess = false
                });

                dynamic? doc = null;
                try
                {
                    // Word belgesini aç
                    doc = wordApp.Documents.Open(
                        FileName: file.FullName,
                        ConfirmConversions: false,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        Visible: false
                    );

                    // 17 = wdExportFormatPDF, 0 = wdExportOptimizeForPrint
                    doc.ExportAsFixedFormat(
                        OutputFileName: outputPdfPath,
                        ExportFormat: 17,
                        OpenAfterExport: false,
                        OptimizeFor: 0
                    );

                    progress.Report(new ConversionProgress
                    {
                        CurrentIndex = i + 1,
                        TotalCount = files.Count,
                        CurrentFileName = file.Name,
                        Status = "Tamamlandı",
                        IsSuccess = true
                    });
                }
                catch (Exception ex)
                {
                    progress.Report(new ConversionProgress
                    {
                        CurrentIndex = i + 1,
                        TotalCount = files.Count,
                        CurrentFileName = file.Name,
                        Status = "Hata",
                        IsSuccess = false,
                        ErrorMessage = ex.Message
                    });
                }
                finally
                {
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(0); // wdDoNotSaveChanges = 0
                        }
                        catch { }
                        Marshal.ReleaseComObject(doc);
                    }
                }
            }
        }
        finally
        {
            try
            {
                wordApp.Quit(0); // wdDoNotSaveChanges = 0
            }
            catch { }
            Marshal.ReleaseComObject(wordApp);
        }
    }

    private async Task ConvertWithLibreOfficeAsync(
        List<FileInfo> files,
        string targetFolder,
        string libreOfficePath,
        IProgress<ConversionProgress> progress,
        CancellationToken cancellationToken)
    {
        for (int i = 0; i < files.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var file = files[i];
            progress.Report(new ConversionProgress
            {
                CurrentIndex = i + 1,
                TotalCount = files.Count,
                CurrentFileName = file.Name,
                Status = "Dönüştürülüyor...",
                IsSuccess = false
            });

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = libreOfficePath,
                    Arguments = $"--headless --convert-to pdf \"{file.FullName}\" --outdir \"{targetFolder}\"",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

                using var process = Process.Start(psi);
                if (process != null)
                {
                    await process.WaitForExitAsync(cancellationToken);
                    if (process.ExitCode == 0)
                    {
                        progress.Report(new ConversionProgress
                        {
                            CurrentIndex = i + 1,
                            TotalCount = files.Count,
                            CurrentFileName = file.Name,
                            Status = "Tamamlandı",
                            IsSuccess = true
                        });
                    }
                    else
                    {
                        string err = await process.StandardError.ReadToEndAsync(cancellationToken);
                        progress.Report(new ConversionProgress
                        {
                            CurrentIndex = i + 1,
                            TotalCount = files.Count,
                            CurrentFileName = file.Name,
                            Status = "Hata",
                            IsSuccess = false,
                            ErrorMessage = string.IsNullOrWhiteSpace(err) ? $"Çıkış kodu: {process.ExitCode}" : err
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                progress.Report(new ConversionProgress
                {
                    CurrentIndex = i + 1,
                    TotalCount = files.Count,
                    CurrentFileName = file.Name,
                    Status = "Hata",
                    IsSuccess = false,
                    ErrorMessage = ex.Message
                });
            }
        }
    }

    private static string? FindLibreOffice()
    {
        string[] commonPaths =
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "program", "soffice.exe")
        };

        foreach (var path in commonPaths)
        {
            if (File.Exists(path)) return path;
        }

        string? envPath = Environment.GetEnvironmentVariable("PATH");
        if (!string.IsNullOrEmpty(envPath))
        {
            foreach (var part in envPath.Split(Path.PathSeparator))
            {
                try
                {
                    string candidate = Path.Combine(part.Trim(), "soffice.exe");
                    if (File.Exists(candidate)) return candidate;
                }
                catch { }
            }
        }

        return null;
    }
}
