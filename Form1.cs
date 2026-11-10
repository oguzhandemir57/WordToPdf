using System.Diagnostics;

namespace WordToPdfApp;

public partial class Form1 : Form
{
    private readonly WordToPdfConverterService _converterService = new();
    private List<FileInfo> _currentWordFiles = new();
    private CancellationTokenSource? _cancellationTokenSource;
    private bool _isConverting = false;

    public Form1()
    {
        InitializeComponent();
        CheckEngineStatus();
    }

    private void CheckEngineStatus()
    {
        var (engine, desc, path) = WordToPdfConverterService.DetectEngine();

        if (engine == ConversionEngine.MicrosoftWord)
        {
            lblEngineBadge.Text = "✔ Microsoft Word Aktif";
            lblEngineBadge.BackColor = Color.FromArgb(16, 185, 129); // Yeşil
            lblEngineBadge.ForeColor = Color.White;
        }
        else if (engine == ConversionEngine.LibreOffice)
        {
            lblEngineBadge.Text = "✔ LibreOffice Aktif";
            lblEngineBadge.BackColor = Color.FromArgb(59, 130, 246); // Mavi
            lblEngineBadge.ForeColor = Color.White;
        }
        else
        {
            lblEngineBadge.Text = "⚠ Word / LibreOffice Bulunamadı";
            lblEngineBadge.BackColor = Color.FromArgb(239, 68, 68); // Kırmızı
            lblEngineBadge.ForeColor = Color.White;
        }
    }

    private void BtnBrowseFolder_Click(object? sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog();
        dialog.Description = "Word dosyalarının bulunduğu klasörü seçin";
        dialog.UseDescriptionForTitle = true;
        dialog.ShowNewFolderButton = false;

        string current = CleanPath(txtFolderPath.Text);
        if (!string.IsNullOrWhiteSpace(current) && Directory.Exists(current))
        {
            dialog.SelectedPath = current;
        }

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            txtFolderPath.Text = dialog.SelectedPath;
            LoadFilesFromFolder(dialog.SelectedPath);
        }
    }

    private void BtnBrowseFiles_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog();
        dialog.Title = "Dönüştürmek İstediğiniz Word Dosyalarını Seçin";
        dialog.Filter = "Word Belgeleri (*.docx;*.doc;*.docm)|*.docx;*.doc;*.docm|Tüm Dosyalar (*.*)|*.*";
        dialog.Multiselect = true;

        if (dialog.ShowDialog(this) == DialogResult.OK && dialog.FileNames.Length > 0)
        {
            txtFolderPath.Text = Path.GetDirectoryName(dialog.FileNames[0]) ?? string.Empty;
            LoadSpecificFiles(dialog.FileNames);
        }
    }

    private void TxtFolderPath_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            e.SuppressKeyPress = true;
            BtnRefreshList_Click(sender, EventArgs.Empty);
        }
    }

    private void BtnRefreshList_Click(object? sender, EventArgs e)
    {
        string rawPath = CleanPath(txtFolderPath.Text);
        if (string.IsNullOrWhiteSpace(rawPath))
        {
            MessageBox.Show(
                "Lütfen bir klasör yolu girin veya '📁 Klasör Seç' / '📄 Word Dosyası Seç' butonlarına tıklayın.",
                "Bilgi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        if (Directory.Exists(rawPath))
        {
            LoadFilesFromFolder(rawPath);
        }
        else if (File.Exists(rawPath))
        {
            LoadSpecificFiles(new[] { rawPath });
        }
        else
        {
            MessageBox.Show(
                $"Belirtilen yol bulunamadı:\n{rawPath}\n\nLütfen klasör yolunun doğruluğundan emin olun.",
                "Yol Bulunamadı",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private void ChkIncludeSubfolders_CheckedChanged(object? sender, EventArgs e)
    {
        string path = CleanPath(txtFolderPath.Text);
        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
        {
            LoadFilesFromFolder(path);
        }
    }

    private void LoadFilesFromFolder(string folderPath)
    {
        try
        {
            Cursor = Cursors.WaitCursor;
            bool includeSub = chkIncludeSubfolders.Checked;
            _currentWordFiles = WordToPdfConverterService.ScanWordFiles(folderPath, includeSub);
            PopulateListView(_currentWordFiles);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Klasör taranırken bir hata oluştu:\n{ex.Message}",
                "Hata",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private void LoadSpecificFiles(string[] filePaths)
    {
        try
        {
            Cursor = Cursors.WaitCursor;
            _currentWordFiles = filePaths
                .Where(f => File.Exists(f))
                .Select(f => new FileInfo(f))
                .Where(f => !f.Name.StartsWith("~$", StringComparison.OrdinalIgnoreCase) &&
                           (f.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
                            f.Extension.Equals(".doc", StringComparison.OrdinalIgnoreCase) ||
                            f.Extension.Equals(".docm", StringComparison.OrdinalIgnoreCase)))
                .OrderBy(f => f.Name)
                .ToList();

            PopulateListView(_currentWordFiles);
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private void PopulateListView(List<FileInfo> files)
    {
        lvFiles.BeginUpdate();
        lvFiles.Items.Clear();

        int index = 1;
        foreach (var file in files)
        {
            var lvi = new ListViewItem(index.ToString());
            lvi.SubItems.Add(file.Name);
            lvi.SubItems.Add(FormatFileSize(file.Length));
            lvi.SubItems.Add("Bekliyor");
            lvi.SubItems.Add(file.DirectoryName ?? string.Empty);
            lvi.Tag = file.FullName;
            lvFiles.Items.Add(lvi);
            index++;
        }
        lvFiles.EndUpdate();

        lblListTitle.Text = $"Bulunan Word Belgeleri ({files.Count} dosya)";
        lblProgressStatus.Text = files.Count > 0
            ? $"{files.Count} adet Word belgesi seçildi. 'Dönüştürmeyi Başlat' butonuna tıklayabilirsiniz."
            : "Seçilen konumda Word belgesi (.docx, .doc) bulunamadı.";

        progressBar.Value = 0;
        lblProgressCount.Text = $"0 / {files.Count} (%0)";
        btnStart.Enabled = files.Count > 0;
    }

    private async void BtnStart_Click(object? sender, EventArgs e)
    {
        if (_isConverting) return;

        if (_currentWordFiles.Count == 0)
        {
            MessageBox.Show(
                "Lütfen önce '📁 Klasör Seç' veya '📄 Word Dosyası Seç' butonları ile dönüştürülecek dosyaları belirleyin.",
                "Uyarı",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var (engine, desc, _) = WordToPdfConverterService.DetectEngine();
        if (engine == ConversionEngine.None)
        {
            var result = MessageBox.Show(
                "Word belgelerini PDF'e dönüştürebilmek için bilgisayarınızda Microsoft Word veya LibreOffice kurulu olmalıdır.\n\n" +
                "Otomatik motor algılanamadı. Yine de dönüştürmeyi denemek istiyor musunuz?",
                "Dönüştürücü Motor Uyarısı",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
            {
                return;
            }
        }

        string targetFolder = WordToPdfConverterService.GetDesktopPdfFolderPath();

        SetConversionState(true);
        _cancellationTokenSource = new CancellationTokenSource();

        progressBar.Minimum = 0;
        progressBar.Maximum = _currentWordFiles.Count;
        progressBar.Value = 0;

        int successCount = 0;
        int errorCount = 0;

        var progress = new Progress<ConversionProgress>(p =>
        {
            if (p.CurrentIndex <= lvFiles.Items.Count)
            {
                var lvi = lvFiles.Items[p.CurrentIndex - 1];
                lvi.SubItems[3].Text = p.Status;

                if (p.IsSuccess)
                {
                    lvi.ForeColor = Color.FromArgb(22, 101, 52); // Koyu Yeşil
                    lvi.SubItems[4].Text = "PDF oluşturuldu";
                }
                else if (!string.IsNullOrEmpty(p.ErrorMessage))
                {
                    lvi.ForeColor = Color.FromArgb(185, 28, 28); // Kırmızı
                    lvi.SubItems[4].Text = p.ErrorMessage;
                }

                lvi.EnsureVisible();
            }

            if (p.Status == "Dönüştürülüyor...")
            {
                lblProgressStatus.Text = $"Dönüştürülüyor: {p.CurrentFileName} ({p.CurrentIndex} / {p.TotalCount})...";
            }
            else
            {
                progressBar.Value = Math.Min(p.CurrentIndex, progressBar.Maximum);
                int percent = (int)((double)progressBar.Value / p.TotalCount * 100);
                lblProgressCount.Text = $"{progressBar.Value} / {p.TotalCount} (%{percent})";

                if (p.IsSuccess) successCount++;
                else if (!string.IsNullOrEmpty(p.ErrorMessage)) errorCount++;
            }
        });

        try
        {
            await _converterService.ConvertFilesAsync(
                _currentWordFiles,
                targetFolder,
                progress,
                _cancellationTokenSource.Token);

            lblProgressStatus.Text = $"İşlem tamamlandı! {successCount} dosya dönüştürüldü" +
                                     (errorCount > 0 ? $", {errorCount} hata oluştu." : ".");

            var openPrompt = MessageBox.Show(
                $"Word belgelerinin PDF'e dönüştürülmesi tamamlandı!\n\n" +
                $"✔ Başarılı: {successCount} dosya\n" +
                (errorCount > 0 ? $"✖ Hatalı: {errorCount} dosya\n" : "") +
                $"\nHedef Klasör:\n{targetFolder}\n\nMasaüstündeki 'PdfDosyaları' klasörünü şimdi açmak ister misiniz?",
                "İşlem Tamamlandı",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);

            if (openPrompt == DialogResult.Yes)
            {
                OpenFolder(targetFolder);
            }
        }
        catch (OperationCanceledException)
        {
            lblProgressStatus.Text = "Dönüştürme işlemi iptal edildi.";
            MessageBox.Show(
                "İşlem kullanıcı tarafından iptal edildi.",
                "Bilgi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            lblProgressStatus.Text = "Hata oluştu!";
            MessageBox.Show(
                $"Dönüştürme sırasında bir hata meydana geldi:\n\n{ex.Message}",
                "Dönüştürme Hatası",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            SetConversionState(false);
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    private void BtnCancel_Click(object? sender, EventArgs e)
    {
        if (_isConverting && _cancellationTokenSource != null)
        {
            btnCancel.Enabled = false;
            btnCancel.Text = "İptal ediliyor...";
            _cancellationTokenSource.Cancel();
        }
    }

    private void BtnOpenFolder_Click(object? sender, EventArgs e)
    {
        string targetFolder = WordToPdfConverterService.GetDesktopPdfFolderPath();
        OpenFolder(targetFolder);
    }

    private void OpenFolder(string path)
    {
        try
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Klasör açılırken hata oluştu:\n{ex.Message}",
                "Hata",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void SetConversionState(bool converting)
    {
        _isConverting = converting;
        btnStart.Enabled = !converting;
        btnSelectFolder.Enabled = !converting;
        btnSelectFiles.Enabled = !converting;
        btnRefreshList.Enabled = !converting;
        chkIncludeSubfolders.Enabled = !converting;
        btnCancel.Enabled = converting;
        btnCancel.Text = "⏹ İptal Et";
        Cursor = converting ? Cursors.WaitCursor : Cursors.Default;
    }

    private void Form1_DragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effect = DragDropEffects.Copy;
        }
        else
        {
            e.Effect = DragDropEffects.None;
        }
    }

    private void Form1_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[]? items = (string[]?)e.Data.GetData(DataFormats.FileDrop);
            if (items != null && items.Length > 0)
            {
                if (Directory.Exists(items[0]))
                {
                    txtFolderPath.Text = items[0];
                    LoadFilesFromFolder(items[0]);
                }
                else
                {
                    txtFolderPath.Text = Path.GetDirectoryName(items[0]) ?? string.Empty;
                    LoadSpecificFiles(items);
                }
            }
        }
    }

    private static string CleanPath(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        return input.Trim().Trim('\"', '\'');
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }
        return $"{len:0.#} {sizes[order]}";
    }
}
