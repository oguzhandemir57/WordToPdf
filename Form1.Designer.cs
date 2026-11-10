namespace WordToPdfApp;

partial class Form1
{
    private System.ComponentModel.IContainer components = null;

    private Panel panelHeader;
    private Label lblTitle;
    private Label lblSubtitle;
    private Label lblEngineBadge;

    private Panel panelFolderSelection;
    private Label lblFolderPrompt;
    private TextBox txtFolderPath;
    private Button btnBrowse;
    private CheckBox chkIncludeSubfolders;

    private Panel panelTargetInfo;
    private Label lblTargetInfo;
    private Button btnOpenTargetFolderTop;

    private Panel panelListContainer;
    private Label lblListTitle;
    private ListView lvFiles;
    private ColumnHeader colIndex;
    private ColumnHeader colName;
    private ColumnHeader colSize;
    private ColumnHeader colStatus;
    private ColumnHeader colMessage;

    private Panel panelBottom;
    private ProgressBar progressBar;
    private Label lblProgressStatus;
    private Label lblProgressCount;
    private Button btnStart;
    private Button btnCancel;
    private Button btnOpenFolder;

    private FolderBrowserDialog folderBrowserDialog;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        panelHeader = new Panel();
        lblTitle = new Label();
        lblSubtitle = new Label();
        lblEngineBadge = new Label();

        panelFolderSelection = new Panel();
        lblFolderPrompt = new Label();
        txtFolderPath = new TextBox();
        btnBrowse = new Button();
        chkIncludeSubfolders = new CheckBox();

        panelTargetInfo = new Panel();
        lblTargetInfo = new Label();
        btnOpenTargetFolderTop = new Button();

        panelListContainer = new Panel();
        lblListTitle = new Label();
        lvFiles = new ListView();
        colIndex = new ColumnHeader();
        colName = new ColumnHeader();
        colSize = new ColumnHeader();
        colStatus = new ColumnHeader();
        colMessage = new ColumnHeader();

        panelBottom = new Panel();
        progressBar = new ProgressBar();
        lblProgressStatus = new Label();
        lblProgressCount = new Label();
        btnStart = new Button();
        btnCancel = new Button();
        btnOpenFolder = new Button();

        folderBrowserDialog = new FolderBrowserDialog();

        // 
        // Form1
        // 
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(860, 640);
        MinimumSize = new Size(720, 500);
        Text = "Word to PDF Dönüştürücü - Masaüstü PdfDosyaları";
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(248, 250, 252);
        AllowDrop = true;

        // 
        // panelHeader
        // 
        panelHeader.Dock = DockStyle.Top;
        panelHeader.Height = 85;
        panelHeader.BackColor = Color.FromArgb(30, 41, 59);
        panelHeader.Padding = new Padding(20, 15, 20, 15);

        lblTitle.Text = "Word to PDF Dönüştürücü";
        lblTitle.Font = new Font("Segoe UI Semibold", 15F, FontStyle.Bold, GraphicsUnit.Point);
        lblTitle.ForeColor = Color.White;
        lblTitle.AutoSize = true;
        lblTitle.Location = new Point(20, 14);

        lblSubtitle.Text = "Seçilen klasördeki Word (.docx, .doc) belgelerini Masaüstünde 'PdfDosyaları' klasörüne dönüştürür.";
        lblSubtitle.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        lblSubtitle.ForeColor = Color.FromArgb(203, 213, 225);
        lblSubtitle.AutoSize = true;
        lblSubtitle.Location = new Point(22, 47);

        lblEngineBadge.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        lblEngineBadge.AutoSize = true;
        lblEngineBadge.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold, GraphicsUnit.Point);
        lblEngineBadge.ForeColor = Color.FromArgb(226, 232, 240);
        lblEngineBadge.BackColor = Color.FromArgb(51, 65, 85);
        lblEngineBadge.Padding = new Padding(8, 4, 8, 4);
        lblEngineBadge.Location = new Point(600, 28);
        lblEngineBadge.Text = "Motor Denetleniyor...";

        panelHeader.Controls.Add(lblTitle);
        panelHeader.Controls.Add(lblSubtitle);
        panelHeader.Controls.Add(lblEngineBadge);

        // 
        // panelFolderSelection
        // 
        panelFolderSelection.Dock = DockStyle.Top;
        panelFolderSelection.Height = 78;
        panelFolderSelection.BackColor = Color.White;
        panelFolderSelection.Padding = new Padding(20, 10, 20, 10);

        lblFolderPrompt.Text = "Word Belgelerinin Bulunduğu Kaynak Klasör:";
        lblFolderPrompt.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point);
        lblFolderPrompt.ForeColor = Color.FromArgb(15, 23, 42);
        lblFolderPrompt.AutoSize = true;
        lblFolderPrompt.Location = new Point(20, 10);

        txtFolderPath.Location = new Point(20, 36);
        txtFolderPath.Size = new Size(580, 29);
        txtFolderPath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        txtFolderPath.PlaceholderText = "Bir klasör yolu yazın veya 'Klasör Seç' butonuna tıklayın (Sürükle-bırak desteklenir)...";

        btnBrowse.Text = "📁 Klasör Seç...";
        btnBrowse.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold, GraphicsUnit.Point);
        btnBrowse.Location = new Point(610, 34);
        btnBrowse.Size = new Size(115, 32);
        btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnBrowse.BackColor = Color.FromArgb(241, 245, 249);
        btnBrowse.FlatStyle = FlatStyle.System;
        btnBrowse.Click += BtnBrowse_Click;

        chkIncludeSubfolders.Text = "Alt klasörleri de tara";
        chkIncludeSubfolders.AutoSize = true;
        chkIncludeSubfolders.Location = new Point(735, 39);
        chkIncludeSubfolders.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        chkIncludeSubfolders.CheckedChanged += ChkIncludeSubfolders_CheckedChanged;

        panelFolderSelection.Controls.Add(lblFolderPrompt);
        panelFolderSelection.Controls.Add(txtFolderPath);
        panelFolderSelection.Controls.Add(btnBrowse);
        panelFolderSelection.Controls.Add(chkIncludeSubfolders);

        // 
        // panelTargetInfo
        // 
        panelTargetInfo.Dock = DockStyle.Top;
        panelTargetInfo.Height = 44;
        panelTargetInfo.BackColor = Color.FromArgb(238, 242, 255);
        panelTargetInfo.Padding = new Padding(20, 8, 20, 8);

        lblTargetInfo.Text = "📂 Hedef Konum: Masaüstü\\PdfDosyaları (Dosya isimleri aynı korunacaktır)";
        lblTargetInfo.ForeColor = Color.FromArgb(67, 56, 202);
        lblTargetInfo.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold, GraphicsUnit.Point);
        lblTargetInfo.AutoSize = true;
        lblTargetInfo.Location = new Point(20, 12);

        btnOpenTargetFolderTop.Text = "Masaüstü Klasörünü Aç";
        btnOpenTargetFolderTop.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point);
        btnOpenTargetFolderTop.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnOpenTargetFolderTop.Location = new Point(680, 8);
        btnOpenTargetFolderTop.Size = new Size(160, 28);
        btnOpenTargetFolderTop.BackColor = Color.White;
        btnOpenTargetFolderTop.FlatStyle = FlatStyle.System;
        btnOpenTargetFolderTop.Click += BtnOpenFolder_Click;

        panelTargetInfo.Controls.Add(lblTargetInfo);
        panelTargetInfo.Controls.Add(btnOpenTargetFolderTop);

        // 
        // panelListContainer
        // 
        panelListContainer.Dock = DockStyle.Fill;
        panelListContainer.Padding = new Padding(20, 10, 20, 10);
        panelListContainer.BackColor = Color.FromArgb(248, 250, 252);

        lblListTitle.Text = "Bulunan Word Belgeleri (0 dosya)";
        lblListTitle.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point);
        lblListTitle.ForeColor = Color.FromArgb(51, 65, 85);
        lblListTitle.Dock = DockStyle.Top;
        lblListTitle.Height = 26;

        lvFiles.Dock = DockStyle.Fill;
        lvFiles.View = View.Details;
        lvFiles.FullRowSelect = true;
        lvFiles.GridLines = true;
        lvFiles.BorderStyle = BorderStyle.FixedSingle;
        lvFiles.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

        colIndex.Text = "#";
        colIndex.Width = 45;

        colName.Text = "Dosya Adı";
        colName.Width = 320;

        colSize.Text = "Boyut";
        colSize.Width = 95;

        colStatus.Text = "Durum";
        colStatus.Width = 130;

        colMessage.Text = "Bilgi / Hata Detayı";
        colMessage.Width = 240;

        lvFiles.Columns.AddRange(new ColumnHeader[] { colIndex, colName, colSize, colStatus, colMessage });

        panelListContainer.Controls.Add(lvFiles);
        panelListContainer.Controls.Add(lblListTitle);

        // 
        // panelBottom
        // 
        panelBottom.Dock = DockStyle.Bottom;
        panelBottom.Height = 115;
        panelBottom.BackColor = Color.White;
        panelBottom.Padding = new Padding(20, 12, 20, 12);

        progressBar.Location = new Point(20, 14);
        progressBar.Size = new Size(820, 22);
        progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

        lblProgressStatus.Text = "Hazır - Klasör seçtikten sonra 'Dönüştürmeyi Başlat' butonuna tıklayın.";
        lblProgressStatus.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        lblProgressStatus.ForeColor = Color.FromArgb(71, 85, 105);
        lblProgressStatus.Location = new Point(20, 42);
        lblProgressStatus.AutoSize = true;

        lblProgressCount.Text = "0 / 0 (%0)";
        lblProgressCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold, GraphicsUnit.Point);
        lblProgressCount.ForeColor = Color.FromArgb(71, 85, 105);
        lblProgressCount.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        lblProgressCount.Location = new Point(730, 42);
        lblProgressCount.TextAlign = ContentAlignment.TopRight;
        lblProgressCount.AutoSize = true;

        btnStart.Text = "▶ Dönüştürmeyi Başlat";
        btnStart.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold, GraphicsUnit.Point);
        btnStart.Size = new Size(200, 38);
        btnStart.Location = new Point(20, 68);
        btnStart.BackColor = Color.FromArgb(16, 185, 129);
        btnStart.ForeColor = Color.White;
        btnStart.FlatStyle = FlatStyle.Flat;
        btnStart.FlatAppearance.BorderSize = 0;
        btnStart.Cursor = Cursors.Hand;
        btnStart.Click += BtnStart_Click;

        btnCancel.Text = "⏹ İptal Et";
        btnCancel.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point);
        btnCancel.Size = new Size(110, 38);
        btnCancel.Location = new Point(230, 68);
        btnCancel.BackColor = Color.FromArgb(239, 68, 68);
        btnCancel.ForeColor = Color.White;
        btnCancel.FlatStyle = FlatStyle.Flat;
        btnCancel.FlatAppearance.BorderSize = 0;
        btnCancel.Cursor = Cursors.Hand;
        btnCancel.Enabled = false;
        btnCancel.Click += BtnCancel_Click;

        btnOpenFolder.Text = "📂 PdfDosyaları Klasörünü Aç";
        btnOpenFolder.Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        btnOpenFolder.Size = new Size(220, 38);
        btnOpenFolder.Location = new Point(620, 68);
        btnOpenFolder.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        btnOpenFolder.BackColor = Color.FromArgb(241, 245, 249);
        btnOpenFolder.ForeColor = Color.FromArgb(30, 41, 59);
        btnOpenFolder.FlatStyle = FlatStyle.Flat;
        btnOpenFolder.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        btnOpenFolder.Cursor = Cursors.Hand;
        btnOpenFolder.Click += BtnOpenFolder_Click;

        panelBottom.Controls.Add(progressBar);
        panelBottom.Controls.Add(lblProgressStatus);
        panelBottom.Controls.Add(lblProgressCount);
        panelBottom.Controls.Add(btnStart);
        panelBottom.Controls.Add(btnCancel);
        panelBottom.Controls.Add(btnOpenFolder);

        // 
        // Add all panels to form
        // 
        Controls.Add(panelListContainer);
        Controls.Add(panelBottom);
        Controls.Add(panelTargetInfo);
        Controls.Add(panelFolderSelection);
        Controls.Add(panelHeader);

        DragEnter += Form1_DragEnter;
        DragDrop += Form1_DragDrop;
    }
}
