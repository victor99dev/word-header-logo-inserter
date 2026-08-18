namespace Cabecalho.LogoInserter.Views;

public sealed class MainForm : Form, IMainView
{
    private readonly TextBox _txtDocumento = new();
    private readonly TextBox _txtLogo = new();
    private readonly TextBox _txtSaida = new();
    private readonly Button _btnDocumento = new();
    private readonly Button _btnLogo = new();
    private readonly Button _btnSaida = new();
    private readonly Button _btnProcessar = new();
    private readonly CheckBox _chkAbrir = new();
    private readonly Label _lblStatus = new();

    public MainForm()
    {
        Text = "Métrica Topo - Inserção de Logo";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(820, 420);
        Size = new Size(930, 480);
        Font = new Font("Segoe UI", 10F);

        BuildLayout();
    }

    public string DocumentoPath
    {
        get => _txtDocumento.Text.Trim();
        set => _txtDocumento.Text = value;
    }

    public string LogoPath
    {
        get => _txtLogo.Text.Trim();
        set => _txtLogo.Text = value;
    }

    public string SaidaPath
    {
        get => _txtSaida.Text.Trim();
        set => _txtSaida.Text = value;
    }

    public bool AbrirAoFinal => _chkAbrir.Checked;

    public event EventHandler? SelecionarDocumento;
    public event EventHandler? SelecionarLogo;
    public event EventHandler? SelecionarSaida;
    public event EventHandler? Processar;

    public string? EscolherDocumento()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Selecione o documento",
            Filter = "Documentos (*.docx;*.docm;*.doc)|*.docx;*.docm;*.doc|Todos os arquivos (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
    }

    public string? EscolherLogo()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Selecione a logomarca",
            Filter = "Imagens (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|Todos os arquivos (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
    }

    public string? EscolherSaida(string sugestaoInicial)
    {
        var nomeArquivo = Path.GetFileName(sugestaoInicial);
        var diretorio = Path.GetDirectoryName(sugestaoInicial);

        using var dialog = new SaveFileDialog
        {
            Title = "Selecione onde salvar o documento processado",
            Filter = "Documento (*.docx)|*.docx|Documento com macro (*.docm)|*.docm|Documento 97-2003 (*.doc)|*.doc",
            FileName = nomeArquivo,
            InitialDirectory = Directory.Exists(diretorio) ? diretorio : null,
            AddExtension = true,
            OverwritePrompt = true
        };

        return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
    }

    public void SetBusy(bool busy)
    {
        UseWaitCursor = busy;
        _btnDocumento.Enabled = !busy;
        _btnLogo.Enabled = !busy;
        _btnSaida.Enabled = !busy;
        _btnProcessar.Enabled = !busy;
        _txtDocumento.Enabled = !busy;
        _txtLogo.Enabled = !busy;
        _txtSaida.Enabled = !busy;
        _chkAbrir.Enabled = !busy;
        Application.DoEvents();
    }

    public void SetStatus(string mensagem) => _lblStatus.Text = mensagem;

    public void MostrarSucesso(string mensagem) =>
        MessageBox.Show(this, mensagem, "Concluído", MessageBoxButtons.OK, MessageBoxIcon.Information);

    public void MostrarErro(string mensagem) =>
        MessageBox.Show(this, mensagem, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);

    private void BuildLayout()
    {
        var main = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(28),
            ColumnCount = 3,
            RowCount = 7
        };

        main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        main.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        main.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var titulo = new Label
        {
            Text = "Inserção automática de logomarca no cabeçalho do documento",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold),
            Anchor = AnchorStyles.Left
        };
        main.Controls.Add(titulo, 0, 0);
        main.SetColumnSpan(titulo, 3);

        AddPathRow(main, 1, "Documento", _txtDocumento, _btnDocumento, "Selecionar...");
        AddPathRow(main, 2, "Logomarca", _txtLogo, _btnLogo, "Selecionar...");
        AddPathRow(main, 3, "Salvar como", _txtSaida, _btnSaida, "Selecionar...");

        _chkAbrir.Text = "Abrir o documento ao finalizar";
        _chkAbrir.Checked = true;
        _chkAbrir.AutoSize = true;
        _chkAbrir.Anchor = AnchorStyles.Left;
        main.Controls.Add(_chkAbrir, 1, 4);
        main.SetColumnSpan(_chkAbrir, 2);

        _btnProcessar.Text = "Inserir logo";
        _btnProcessar.AutoSize = true;
        _btnProcessar.Padding = new Padding(18, 7, 18, 7);
        _btnProcessar.Anchor = AnchorStyles.Left;
        main.Controls.Add(_btnProcessar, 1, 5);

        _lblStatus.Text = "Selecione o documento, a logo e o arquivo de saída.";
        _lblStatus.AutoSize = true;
        _lblStatus.Anchor = AnchorStyles.Left | AnchorStyles.Top;
        main.Controls.Add(_lblStatus, 0, 6);
        main.SetColumnSpan(_lblStatus, 3);

        Controls.Add(main);

        _btnDocumento.Click += (_, _) => SelecionarDocumento?.Invoke(this, EventArgs.Empty);
        _btnLogo.Click += (_, _) => SelecionarLogo?.Invoke(this, EventArgs.Empty);
        _btnSaida.Click += (_, _) => SelecionarSaida?.Invoke(this, EventArgs.Empty);
        _btnProcessar.Click += (_, _) => Processar?.Invoke(this, EventArgs.Empty);
    }

    private void InitializeComponent()
    {

    }

    private static void AddPathRow(
        TableLayoutPanel layout,
        int row,
        string labelText,
        TextBox textBox,
        Button button,
        string buttonText)
    {
        var label = new Label
        {
            Text = labelText,
            AutoSize = true,
            Anchor = AnchorStyles.Left
        };

        textBox.Dock = DockStyle.Fill;
        textBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;

        button.Text = buttonText;
        button.AutoSize = true;
        button.Anchor = AnchorStyles.Right;

        layout.Controls.Add(label, 0, row);
        layout.Controls.Add(textBox, 1, row);
        layout.Controls.Add(button, 2, row);
    }
}
