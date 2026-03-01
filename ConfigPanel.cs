using NotificadorBajasHitssApp.Config;
using NotificadorBajasHitssApp.Services;

namespace NotificadorBajasHitssApp;

/// <summary>Panel con controles para editar AppConfig (vista de configuración) con selectores de carpeta.</summary>
public class ConfigPanel : UserControl
{
    private readonly Dictionary<string, TextBox> _inputs = new();
    private readonly ToolTip _tooltip;
    private readonly Dictionary<Control, string> _hintTitles = new();
    private ComboBox? _comboOutlookCuenta;

    private const int LabelWidth  = 200;
    private const int InputWidth  = 360;
    private const int ButtonWidth = 36;
    private const int RowHeight   = 32;

    // ── Paleta ─────────────────────────────────────────────
    private static readonly Color SectionHeaderBg  = Color.FromArgb(234, 240, 255);
    private static readonly Color SectionAccent    = Color.FromArgb(0, 120, 212);
    private static readonly Color SectionTitleFore = Color.FromArgb(12, 30, 72);
    private static readonly Color LabelFore        = Color.FromArgb(44, 52, 72);
    private static readonly Color BtnBrowseBorder  = Color.FromArgb(0, 120, 212);
    private static readonly Color BtnBrowseFore    = Color.FromArgb(0, 100, 190);
    private static readonly Color HintBg           = Color.FromArgb(214, 230, 255);
    private static readonly Color HintFore         = Color.FromArgb(0, 88, 196);
    private static readonly Font  LabelFont        = new("Segoe UI", 9f);
    private static readonly Font  InputFont        = new("Segoe UI", 9f);
    private static readonly Font  SectionFont      = new("Segoe UI", 9f, FontStyle.Bold);
    private static readonly Font  HintFont         = new("Segoe UI", 7f, FontStyle.Bold);

    // ── Descripciones de ayuda ─────────────────────────────
    private static readonly Dictionary<string, (string Title, string Desc)> Hints = new()
    {
        [nameof(AppConfig.FolderTemporal)]  = ("Carpeta temporal",
            "Carpeta donde se guardan temporalmente los archivos\n" +
            "descargados del correo antes de ser procesados.\n" +
            "Se limpia automáticamente al inicio de cada ejecución."),

        [nameof(AppConfig.FolderUser)]      = ("Carpeta usuario (Excel)",
            "Carpeta de destino donde se almacena el archivo Excel\n" +
            "con las bajas del día. El archivo se nombra automáticamente\n" +
            "con la fecha en formato dd.MM.yy."),

        [nameof(AppConfig.FolderBASE)]      = ("Carpeta base (CSV)",
            "Carpeta que contiene el archivo CSV con la base de datos\n" +
            "maestra de empleados activos (xFileBase).\n" +
            "Se actualiza al procesar las bajas."),

        [nameof(AppConfig.FolderBCKP)]      = ("Carpeta backup",
            "Carpeta donde se guardan copias de seguridad del archivo\n" +
            "base antes de actualizarlo. El backup incluye\n" +
            "la fecha en el nombre del archivo."),

        [nameof(AppConfig.AsuntoCorreoR)]   = ("Asunto a buscar",
            "Texto fijo que debe contener el asunto del correo\n" +
            "a buscar en la carpeta de Outlook configurada.\n" +
            "Ejemplo: 'CESE DE PERSONAL - '"),

        [nameof(AppConfig.OutlookCuenta)]   = ("Cuenta de Outlook",
            "Si tienes varias cuentas en Outlook, elige en cuál buscar.\n" +
            "Dejar en '(Primera cuenta)' usa la cuenta por defecto.\n" +
            "Pulsa 'Cargar' para listar las cuentas detectadas."),

        [nameof(AppConfig.OutlookCarpeta)]  = ("Carpeta Outlook",
            "Ruta de la carpeta en Outlook donde el proceso\n" +
            "busca el correo con las bajas.\n" +
            "Ejemplo: 'Bandeja de entrada\\C.H_BAJAS'"),

        [nameof(AppConfig.CorreoTo)]        = ("Correo destinatario",
            "Dirección de correo electrónico del destinatario\n" +
            "que recibe las notificaciones de bajas procesadas\n" +
            "y las alertas de error o correo no encontrado."),

        [nameof(AppConfig.AsuntoCorreoS)]   = ("Asunto notificación",
            "Asunto del correo de notificación que se envía\n" +
            "con el resultado del proceso. La fecha del día\n" +
            "se añade automáticamente al final."),

        [nameof(AppConfig.Proceso)]         = ("Nombre del proceso",
            "Nombre identificador de este proceso.\n" +
            "Se utiliza en reportes y en el cuerpo\n" +
            "de las notificaciones automáticas."),

        [nameof(AppConfig.FechaDia)]        = ("Días a restar",
            "Número de días que se restan a la fecha actual\n" +
            "para determinar la fecha de corte de las bajas.\n" +
            "Valor típico: 1 (ayer)."),

        [nameof(AppConfig.SheetName)]       = ("Nombre hoja Excel",
            "Nombre exacto de la hoja del archivo Excel\n" +
            "que contiene los datos de bajas a procesar.\n" +
            "Ejemplo: 'Hoja1'"),

        [nameof(AppConfig.FileBase)]        = ("Archivo base CSV",
            "Nombre del archivo CSV con la base de datos\n" +
            "maestra de empleados (debe incluir extensión).\n" +
            "Ejemplo: 'BASE HITSS.csv'"),

        [nameof(AppConfig.FileBkp)]         = ("Prefijo archivo backup",
            "Prefijo del nombre del archivo de backup.\n" +
            "Se le añade la fecha automáticamente al generarlo.\n" +
            "Ejemplo: 'BASE HITSS BKP'"),
    };

    public ConfigPanel()
    {
        _tooltip = new ToolTip
        {
            ShowAlways   = true,
            InitialDelay = 350,
            AutoPopDelay = 10000,
            ReshowDelay  = 200,
            IsBalloon    = false
        };
        _tooltip.Popup += OnHintPopup;

        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(2);
        BackColor = Color.Transparent;
        MinimumSize = new Size(520, 400);

        var main = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents  = false,
            AutoSize      = true,
            AutoSizeMode  = AutoSizeMode.GrowAndShrink,
            Padding       = new Padding(0),
            BackColor     = Color.Transparent
        };

        // ── Sección: Carpetas ──────────────────────────────
        var (sCarpetas, tCarpetas) = NewSection("  Carpetas");
        AddRow(tCarpetas, "Carpeta temporal",        nameof(AppConfig.FolderTemporal), withFolder: true);
        AddRow(tCarpetas, "Carpeta usuario (Excel)", nameof(AppConfig.FolderUser),     withFolder: true);
        AddRow(tCarpetas, "Carpeta BASE (CSV)",      nameof(AppConfig.FolderBASE),     withFolder: true);
        AddRow(tCarpetas, "Carpeta backup",          nameof(AppConfig.FolderBCKP),     withFolder: true);
        main.Controls.Add(sCarpetas);

        // ── Sección: Correo y Outlook ──────────────────────
        var (sCorreo, tCorreo) = NewSection("  Correo y Outlook");
        AddRow(tCorreo, "Asunto a buscar (texto fijo)", nameof(AppConfig.AsuntoCorreoR));
        AddRowOutlookCuenta(tCorreo);
        AddRow(tCorreo, "Carpeta Outlook",              nameof(AppConfig.OutlookCarpeta));
        AddRow(tCorreo, "Correo destinatario",          nameof(AppConfig.CorreoTo));
        AddRow(tCorreo, "Asunto notificación",          nameof(AppConfig.AsuntoCorreoS));
        main.Controls.Add(sCorreo);

        // ── Sección: Proceso y base ────────────────────────
        var (sProceso, tProceso) = NewSection("  Proceso y base");
        AddRow(tProceso, "Nombre proceso",         nameof(AppConfig.Proceso));
        AddRow(tProceso, "Días a restar (fecha)",  nameof(AppConfig.FechaDia));
        AddRow(tProceso, "Nombre hoja Excel",      nameof(AppConfig.SheetName));
        AddRow(tProceso, "Archivo base CSV",       nameof(AppConfig.FileBase));
        AddRow(tProceso, "Prefijo archivo backup", nameof(AppConfig.FileBkp));
        main.Controls.Add(sProceso);

        Controls.Add(main);
    }

    private void OnHintPopup(object? sender, PopupEventArgs e)
    {
        if (e.AssociatedControl != null && _hintTitles.TryGetValue(e.AssociatedControl, out var title))
            _tooltip.ToolTipTitle = title;
    }

    private static (TableLayoutPanel container, TableLayoutPanel table) NewSection(string title)
    {
        var container = new TableLayoutPanel
        {
            ColumnCount   = 1,
            RowCount      = 2,
            AutoSize      = true,
            AutoSizeMode  = AutoSizeMode.GrowAndShrink,
            Margin        = new Padding(0, 0, 0, 10),
            BackColor     = Color.White,
            CellBorderStyle = TableLayoutPanelCellBorderStyle.None
        };
        container.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        container.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        container.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var header = new Panel { Dock = DockStyle.Fill, BackColor = SectionHeaderBg };
        var accentBar = new Panel { Dock = DockStyle.Left, Width = 3, BackColor = SectionAccent };
        var lblTitle = new Label
        {
            Text     = title,
            Font     = SectionFont,
            ForeColor = SectionTitleFore,
            AutoSize = true,
            Anchor   = AnchorStyles.Left | AnchorStyles.Top,
            Location = new Point(10, 6)
        };
        header.Controls.Add(accentBar);
        header.Controls.Add(lblTitle);
        container.Controls.Add(header, 0, 0);

        var table = new TableLayoutPanel
        {
            AutoSize      = true,
            AutoSizeMode  = AutoSizeMode.GrowAndShrink,
            ColumnCount   = 3,
            Padding       = new Padding(8, 6, 8, 6),
            BackColor     = Color.White
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, LabelWidth));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, InputWidth));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, ButtonWidth + 4));
        container.Controls.Add(table, 0, 1);

        return (container, table);
    }

    private void AddRow(TableLayoutPanel layout, string label, string key, bool withFolder = false)
    {
        int row = layout.RowCount;
        layout.RowCount++;
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, RowHeight));

        // ── Celda de etiqueta: label + icono "?" ──────────
        var cellFlow = new FlowLayoutPanel
        {
            AutoSize      = true,
            WrapContents  = false,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor     = Color.Transparent,
            Anchor        = AnchorStyles.Left,
            Margin        = new Padding(0, 6, 8, 0)
        };

        var lbl = new Label
        {
            Text      = label,
            AutoSize  = true,
            Font      = LabelFont,
            ForeColor = LabelFore,
            Margin    = new Padding(0, 1, 3, 0)
        };
        cellFlow.Controls.Add(lbl);

        if (Hints.TryGetValue(key, out var hint))
        {
            var hintLbl = new Label
            {
                Text      = "?",
                AutoSize  = false,
                Size      = new Size(15, 15),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = HintBg,
                ForeColor = HintFore,
                Font      = HintFont,
                Cursor    = Cursors.Help,
                Margin    = new Padding(2, 1, 0, 0)
            };
            _tooltip.SetToolTip(hintLbl, hint.Desc);
            _hintTitles[hintLbl] = hint.Title;
            cellFlow.Controls.Add(hintLbl);
        }

        layout.Controls.Add(cellFlow, 0, row);

        // ── Celda de entrada ───────────────────────────────
        var tb = new TextBox
        {
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Font   = InputFont,
            Margin = new Padding(0, 4, 4, 0),
            Height = 24
        };
        if (key == nameof(AppConfig.FechaDia)) tb.Text = "1";
        layout.Controls.Add(tb, 1, row);
        _inputs[key] = tb;

        // ── Celda de botón de carpeta ──────────────────────
        if (withFolder)
        {
            var btn = new Button
            {
                Text      = "...",
                Width     = ButtonWidth,
                Height    = 24,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = BtnBrowseFore,
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Margin    = new Padding(0, 4, 0, 0),
                Cursor    = Cursors.Hand
            };
            btn.FlatAppearance.BorderColor         = BtnBrowseBorder;
            btn.FlatAppearance.BorderSize          = 1;
            btn.FlatAppearance.MouseOverBackColor  = Color.FromArgb(224, 236, 255);
            var keyCapture = key;
            btn.Click += (_, _) => SeleccionarCarpeta(keyCapture);
            layout.Controls.Add(btn, 2, row);
        }
        else
        {
            layout.Controls.Add(new Panel(), 2, row);
        }
    }

    private void AddRowOutlookCuenta(TableLayoutPanel layout)
    {
        int row = layout.RowCount;
        layout.RowCount++;
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, RowHeight));

        var cellFlow = new FlowLayoutPanel
        {
            AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight,
            BackColor = Color.Transparent, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0)
        };
        var lbl = new Label { Text = "Cuenta de Outlook", AutoSize = true, Font = LabelFont, ForeColor = LabelFore, Margin = new Padding(0, 1, 3, 0) };
        cellFlow.Controls.Add(lbl);
        if (Hints.TryGetValue(nameof(AppConfig.OutlookCuenta), out var hint))
        {
            var hintLbl = new Label { Text = "?", AutoSize = false, Size = new Size(15, 15), TextAlign = ContentAlignment.MiddleCenter, BackColor = HintBg, ForeColor = HintFore, Font = HintFont, Cursor = Cursors.Help, Margin = new Padding(2, 1, 0, 0) };
            _tooltip.SetToolTip(hintLbl, hint.Desc);
            _hintTitles[hintLbl] = hint.Title;
            cellFlow.Controls.Add(hintLbl);
        }
        layout.Controls.Add(cellFlow, 0, row);

        _comboOutlookCuenta = new ComboBox
        {
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Font = InputFont,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Margin = new Padding(0, 4, 4, 0),
            Height = 24
        };
        _comboOutlookCuenta.Items.Add("(Primera cuenta)");
        _comboOutlookCuenta.SelectedIndex = 0;
        layout.Controls.Add(_comboOutlookCuenta, 1, row);

        var btnCargar = new Button
        {
            Text = "Cargar",
            Width = Math.Max(ButtonWidth + 24, 56),
            Height = 24,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            ForeColor = BtnBrowseFore,
            Font = new Font("Segoe UI", 8.5f),
            Margin = new Padding(0, 4, 0, 0),
            Cursor = Cursors.Hand
        };
        btnCargar.FlatAppearance.BorderColor = BtnBrowseBorder;
        btnCargar.FlatAppearance.BorderSize = 1;
        btnCargar.FlatAppearance.MouseOverBackColor = Color.FromArgb(224, 236, 255);
        btnCargar.Click += (_, _) =>
        {
            var cuentas = OutlookService.ObtenerCarpetasRaiz();
            var sel = _comboOutlookCuenta?.SelectedItem?.ToString();
            _comboOutlookCuenta!.Items.Clear();
            _comboOutlookCuenta.Items.Add("(Primera cuenta)");
            foreach (var c in cuentas)
                _comboOutlookCuenta.Items.Add(c);
            _comboOutlookCuenta.SelectedIndex = 0;
            if (!string.IsNullOrEmpty(sel))
            {
                for (int i = 0; i < _comboOutlookCuenta.Items.Count; i++)
                {
                    if (string.Equals(_comboOutlookCuenta.Items[i]?.ToString(), sel, StringComparison.OrdinalIgnoreCase))
                    { _comboOutlookCuenta.SelectedIndex = i; break; }
                }
            }
        };
        layout.Controls.Add(btnCargar, 2, row);
    }

    private void SeleccionarCarpeta(string key)
    {
        using var dlg = new FolderBrowserDialog
        {
            Description        = "Seleccionar carpeta",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true
        };
        var current = _inputs[key].Text?.Trim() ?? "";
        if (!string.IsNullOrEmpty(current) && Directory.Exists(current))
            dlg.SelectedPath = current;
        if (dlg.ShowDialog() == DialogResult.OK)
            _inputs[key].Text = dlg.SelectedPath;
    }

    /// <summary>Rellena las rutas de carpetas con la estructura estándar bajo <paramref name="rootPath"/>.</summary>
    public void SetFolderStructure(string rootPath)
    {
        _inputs[nameof(AppConfig.FolderTemporal)].Text = Path.Combine(rootPath, "Temporal");
        _inputs[nameof(AppConfig.FolderUser)].Text     = Path.Combine(rootPath, "Usuario");
        _inputs[nameof(AppConfig.FolderBASE)].Text     = Path.Combine(rootPath, "Base");
        _inputs[nameof(AppConfig.FolderBCKP)].Text     = Path.Combine(rootPath, "Backup");
    }

    public void LoadFrom(AppConfig config)
    {
        _inputs[nameof(AppConfig.Proceso)].Text        = config.Proceso ?? "";
        _inputs[nameof(AppConfig.FechaDia)].Text       = config.FechaDia.ToString();
        _inputs[nameof(AppConfig.FolderTemporal)].Text = config.FolderTemporal ?? "";
        _inputs[nameof(AppConfig.FolderUser)].Text     = config.FolderUser ?? "";
        _inputs[nameof(AppConfig.AsuntoCorreoR)].Text  = config.AsuntoCorreoR ?? "";
        _inputs[nameof(AppConfig.OutlookCarpeta)].Text = config.OutlookCarpeta ?? "";
        _inputs[nameof(AppConfig.SheetName)].Text      = config.SheetName ?? "";
        _inputs[nameof(AppConfig.FolderBASE)].Text     = config.FolderBASE ?? "";
        _inputs[nameof(AppConfig.FileBase)].Text       = config.FileBase ?? "";
        _inputs[nameof(AppConfig.FolderBCKP)].Text     = config.FolderBCKP ?? "";
        _inputs[nameof(AppConfig.FileBkp)].Text        = config.FileBkp ?? "";
        _inputs[nameof(AppConfig.CorreoTo)].Text       = config.CorreoTo ?? "";
        _inputs[nameof(AppConfig.AsuntoCorreoS)].Text  = config.AsuntoCorreoS ?? "";

        if (_comboOutlookCuenta != null)
        {
            var cuenta = (config.OutlookCuenta ?? "").Trim();
            if (string.IsNullOrEmpty(cuenta) || cuenta == "(Primera cuenta)")
            {
                _comboOutlookCuenta.SelectedIndex = 0;
            }
            else
            {
                var found = false;
                for (int i = 0; i < _comboOutlookCuenta.Items.Count; i++)
                {
                    if (string.Equals(_comboOutlookCuenta.Items[i]?.ToString(), cuenta, StringComparison.OrdinalIgnoreCase))
                    { _comboOutlookCuenta.SelectedIndex = i; found = true; break; }
                }
                if (!found)
                {
                    _comboOutlookCuenta.Items.Add(cuenta);
                    _comboOutlookCuenta.SelectedItem = cuenta;
                }
            }
        }
    }

    public AppConfig SaveTo()
    {
        int.TryParse(_inputs[nameof(AppConfig.FechaDia)].Text, out int fd);
        var cuentaOutlook = "";
        if (_comboOutlookCuenta?.SelectedItem != null)
        {
            var s = _comboOutlookCuenta.SelectedItem.ToString() ?? "";
            if (!string.IsNullOrEmpty(s) && s != "(Primera cuenta)")
                cuentaOutlook = s.Trim();
        }
        return new AppConfig
        {
            Proceso        = _inputs[nameof(AppConfig.Proceso)].Text?.Trim() ?? "",
            FechaDia       = fd,
            FolderTemporal = _inputs[nameof(AppConfig.FolderTemporal)].Text?.Trim() ?? "",
            FolderUser     = _inputs[nameof(AppConfig.FolderUser)].Text?.Trim() ?? "",
            AsuntoCorreoR  = _inputs[nameof(AppConfig.AsuntoCorreoR)].Text?.Trim() ?? "CESE DE PERSONAL - ",
            OutlookCuenta  = cuentaOutlook,
            OutlookCarpeta = _inputs[nameof(AppConfig.OutlookCarpeta)].Text?.Trim() ?? "Bandeja de entrada\\C.H_BAJAS",
            SheetName      = _inputs[nameof(AppConfig.SheetName)].Text?.Trim() ?? "Hoja1",
            FolderBASE     = _inputs[nameof(AppConfig.FolderBASE)].Text?.Trim() ?? "",
            FileBase       = _inputs[nameof(AppConfig.FileBase)].Text?.Trim() ?? "BASE HITSS.csv",
            FolderBCKP     = _inputs[nameof(AppConfig.FolderBCKP)].Text?.Trim() ?? "",
            FileBkp        = _inputs[nameof(AppConfig.FileBkp)].Text?.Trim() ?? "",
            CorreoTo       = _inputs[nameof(AppConfig.CorreoTo)].Text?.Trim() ?? "",
            AsuntoCorreoS  = _inputs[nameof(AppConfig.AsuntoCorreoS)].Text?.Trim() ?? ""
        };
    }
}
