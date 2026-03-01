using NotificadorBajasHitssApp.Config;
using NotificadorBajasHitssApp.Services;

namespace NotificadorBajasHitssApp;

public class MainForm : Form
{
    private readonly ConfigPanel _configPanel;
    private readonly RichTextBox _logBox;
    private readonly Button _btnRun;
    private readonly Button _btnGuardarConfig;
    private readonly Button _btnCancel;
    private readonly Button _btnCrearCarpetas;
    private CancellationTokenSource? _cts;

    // ── Paleta ─────────────────────────────────────────────
    private static readonly Color HeaderBg       = Color.FromArgb(13, 17, 35);
    private static readonly Color HeaderAccent   = Color.FromArgb(0, 120, 212);
    private static readonly Color SurfaceAlt     = Color.FromArgb(245, 246, 250);
    private static readonly Color BorderColor    = Color.FromArgb(218, 224, 236);
    private static readonly Color TextMain       = Color.FromArgb(22, 27, 44);
    private static readonly Color BtnSaveColor   = Color.FromArgb(0, 120, 212);
    private static readonly Color BtnRunColor    = Color.FromArgb(16, 124, 16);
    private static readonly Color BtnCancelColor = Color.FromArgb(162, 36, 28);
    private static readonly Color BtnFolderColor = Color.FromArgb(90, 70, 160);

    // Paleta del log (terminal oscuro)
    private static readonly Color LogBg        = Color.FromArgb(14, 16, 26);
    private static readonly Color LogDefault   = Color.FromArgb(200, 208, 224);
    private static readonly Color LogTimestamp = Color.FromArgb(76, 92, 120);
    private static readonly Color LogSuccess   = Color.FromArgb(68, 195, 148);
    private static readonly Color LogError     = Color.FromArgb(230, 74, 64);
    private static readonly Color LogWarning   = Color.FromArgb(218, 174, 56);

    public MainForm()
    {
        Text = "Notificador · Bajas de Usuarios Hitss";
        Size = new Size(1060, 720);
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(820, 540);
        WindowState = FormWindowState.Maximized;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9f);

        AsignarIcono();

        // ── Raíz: 3 filas (header | contenido | footer) ──────
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(0),
            Margin = new Padding(0),
            BackColor = Color.Transparent
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));

        // ── HEADER ────────────────────────────────────────────
        var header = new Panel { Dock = DockStyle.Fill, BackColor = HeaderBg };
        var accentStripe = new Panel { Dock = DockStyle.Left, Width = 4, BackColor = HeaderAccent };
        var lblTitle = new Label
        {
            Text = "Notificador de Bajas de Usuarios",
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize = true,
            Location = new Point(22, 11)
        };
        var lblSub = new Label
        {
            Text = $"Hitss  ·  v{UpdateService.CurrentVersion.ToString(3)}  ·  Detección de ceses y envío de notificaciones automáticas",
            Font = new Font("Segoe UI", 8f),
            ForeColor = Color.FromArgb(136, 152, 180),
            AutoSize = true,
            Location = new Point(23, 36)
        };
        header.Controls.Add(accentStripe);
        header.Controls.Add(lblTitle);
        header.Controls.Add(lblSub);

        // ── CONTENIDO: split horizontal (config | log) ────────
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterWidth = 5,
            BackColor = BorderColor
        };
        split.Panel1.BackColor = SurfaceAlt;
        split.Panel2.BackColor = Color.White;

        // Panel izquierdo: Configuración
        var leftWrap = new Panel { Dock = DockStyle.Fill, BackColor = SurfaceAlt };
        var leftBar = new Panel { Dock = DockStyle.Top, Height = 36, BackColor = SurfaceAlt };
        var leftBarBorder = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = BorderColor };
        var lblCfgHdr = new Label
        {
            Text = "⚙   Configuración",
            Font = new Font("Segoe UI", 9.25f, FontStyle.Bold),
            ForeColor = TextMain,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(14, 0, 0, 0)
        };
        leftBar.Controls.Add(leftBarBorder);
        leftBar.Controls.Add(lblCfgHdr);

        var configScroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10, 8, 10, 8) };
        _configPanel = new ConfigPanel();
        configScroll.Controls.Add(_configPanel);
        leftWrap.Controls.Add(configScroll);
        leftWrap.Controls.Add(leftBar);
        split.Panel1.Controls.Add(leftWrap);

        // Panel derecho: Log
        var rightWrap = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
        var rightBar = new Panel { Dock = DockStyle.Top, Height = 36, BackColor = Color.White };
        var rightBarBorder = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = BorderColor };
        var lblLogHdr = new Label
        {
            Text = "▶   Registro de ejecución",
            Font = new Font("Segoe UI", 9.25f, FontStyle.Bold),
            ForeColor = TextMain,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(14, 0, 0, 0)
        };
        rightBar.Controls.Add(rightBarBorder);
        rightBar.Controls.Add(lblLogHdr);

        _logBox = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BackColor = LogBg,
            ForeColor = LogDefault,
            Font = new Font("Consolas", 8.75f),
            BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            WordWrap = true,
            DetectUrls = false
        };
        rightWrap.Controls.Add(_logBox);
        rightWrap.Controls.Add(rightBar);
        split.Panel2.Controls.Add(rightWrap);

        // ── FOOTER ────────────────────────────────────────────
        var footer = new Panel { Dock = DockStyle.Fill, BackColor = SurfaceAlt };
        var footerBorder = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = BorderColor };

        _btnGuardarConfig  = MakeBtn("Guardar configuración", BtnSaveColor,   176, 32);
        _btnRun            = MakeBtn("Ejecutar proceso",      BtnRunColor,    148, 32);
        _btnCancel         = MakeBtn("Cancelar",              BtnCancelColor, 100, 32);
        _btnCrearCarpetas  = MakeBtn("Crear estructura de carpetas", BtnFolderColor, 210, 32);
        _btnCancel.Visible = false;

        var btnFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(14, 8, 0, 0),
            BackColor = Color.Transparent
        };
        btnFlow.Controls.Add(_btnGuardarConfig);
        btnFlow.Controls.Add(_btnRun);
        btnFlow.Controls.Add(_btnCancel);

        // Separador visual entre acciones principales y utilitarios
        var sep = new Panel { Width = 16, Height = 32, BackColor = Color.Transparent };
        btnFlow.Controls.Add(sep);
        btnFlow.Controls.Add(_btnCrearCarpetas);

        footer.Controls.Add(footerBorder);
        footer.Controls.Add(btnFlow);

        // ── Ensamble ──────────────────────────────────────────
        root.Controls.Add(header, 0, 0);
        root.Controls.Add(split,  0, 1);
        root.Controls.Add(footer, 0, 2);
        Controls.Add(root);

        // ── Eventos ───────────────────────────────────────────
        _btnGuardarConfig.Click += (_, _) => GuardarConfig();
        _btnRun.Click           += (_, _) => EjecutarProceso();
        _btnCancel.Click        += (_, _) => _cts?.Cancel();
        _btnCrearCarpetas.Click += (_, _) => CrearEstructuraCarpetas();

        Load += async (_, _) =>
        {
            // Panel1 (config) necesita al menos el ancho del ConfigPanel (520) para no cortar opciones
            split.Panel1MinSize = 520;
            split.Panel2MinSize = 260;
            // Por defecto dar suficiente espacio al panel de configuración para ver todas las opciones
            var distanciaInicial = 560;
            split.SplitterDistance = Math.Clamp(distanciaInicial, split.Panel1MinSize,
                Math.Max(split.Panel1MinSize, Width - split.Panel2MinSize - split.SplitterWidth));

            var cfg = ConfigService.Load();
            _configPanel.LoadFrom(cfg);
            Log("Configuración cargada. Revisa las rutas y el asunto (ej. CESE DE PERSONAL - ).");

            await CheckForUpdatesAsync();
        };
    }

    private static Button MakeBtn(string text, Color back, int w, int h)
    {
        var btn = new Button
        {
            Text = text,
            Size = new Size(w, h),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = back,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 8.75f),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = ControlPaint.Dark(back, 0.12f);
        return btn;
    }

    // ── Auto-update ───────────────────────────────────────────
    private async Task CheckForUpdatesAsync()
    {
        Log("Buscando actualizaciones...");
        var release = await UpdateService.CheckAsync();

        if (!release.HasUpdate)
        {
            Log("La aplicación está actualizada.");
            return;
        }

        Log($"Nueva versión disponible: {release.Tag} (actual: v{UpdateService.CurrentVersion.ToString(3)})");

        var msg = $"Hay una nueva versión disponible: {release.Tag}\n" +
                  $"Versión instalada: v{UpdateService.CurrentVersion.ToString(3)}\n\n" +
                  $"¿Deseas actualizar ahora?";

        var result = MessageBox.Show(msg, "Actualización disponible",
            MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);

        if (result == DialogResult.Yes)
            await UpdateService.ApplyAsync(release, Log);
    }

    // ── Crear estructura de carpetas ──────────────────────────
    private void CrearEstructuraCarpetas()
    {
        // Proponer D:\BotHitss si la unidad existe, si no C:\BotHitss
        var driveDefault = Directory.Exists(@"D:\") ? @"D:\BotHitss" : @"C:\BotHitss";

        using var dlg = new FolderBrowserDialog
        {
            Description        = "Selecciona la carpeta raíz donde se creará la estructura de BotHitss",
            UseDescriptionForTitle = true,
            ShowNewFolderButton = true,
            SelectedPath       = Path.GetDirectoryName(driveDefault)!
        };

        if (dlg.ShowDialog() != DialogResult.OK) return;

        var root = Path.Combine(dlg.SelectedPath, "BotHitss");
        var carpetas = new[]
        {
            Path.Combine(root, "Temporal"),
            Path.Combine(root, "Usuario"),
            Path.Combine(root, "Base"),
            Path.Combine(root, "Backup")
        };

        try
        {
            foreach (var c in carpetas)
                Directory.CreateDirectory(c);

            _configPanel.SetFolderStructure(root);

            Log($"Estructura de carpetas creada en {root}:");
            foreach (var c in carpetas)
                Log($"  ✓ {c}");
            Log("Campos actualizados. Recuerda guardar la configuración.");
        }
        catch (Exception ex)
        {
            Log($"Error al crear carpetas: {ex.Message}");
        }
    }

    // ── Guardar config ────────────────────────────────────────
    private void GuardarConfig()
    {
        var cfg = _configPanel.SaveTo();
        ConfigService.Save(cfg);
        Log("Configuración guardada en config.json.");
    }

    // ── Log coloreado ─────────────────────────────────────────
    private void Log(string message)
    {
        if (_logBox.IsDisposed) return;

        var isError   = message.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0;
        var isSuccess = !isError && (message.Contains("finalizado")    || message.Contains("guardada")   ||
                                     message.Contains("movido")        || message.Contains("cargada")    ||
                                     message.Contains("limpiada")      || message.Contains("creada")     ||
                                     message.Contains("actualizada")   || message.Contains("✓"));
        var isWarning = !isError && !isSuccess && (message.Contains("cancelado")    ||
                                                   message.Contains("No se encontró") ||
                                                   message.Contains("actualizaciones") ||
                                                   message.Contains("disponible"));
        var msgColor = isError ? LogError : isSuccess ? LogSuccess : isWarning ? LogWarning : LogDefault;

        void Append()
        {
            _logBox.SelectionStart  = _logBox.TextLength;
            _logBox.SelectionLength = 0;
            _logBox.SelectionColor  = LogTimestamp;
            _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] ");
            _logBox.SelectionColor  = msgColor;
            _logBox.AppendText(message + Environment.NewLine);
            _logBox.ScrollToCaret();
        }

        if (_logBox.InvokeRequired) _logBox.Invoke(Append);
        else Append();
    }

    // ── Ejecutar proceso ──────────────────────────────────────
    private async void EjecutarProceso()
    {
        _btnRun.Enabled    = false;
        _btnCancel.Visible = true;
        _cts = new CancellationTokenSource();
        try
        {
            var config = _configPanel.SaveTo();
            ConfigService.Save(config);

            var asuntoBusqueda = config.AsuntoCorreoR.Trim();
            Log($"Iniciando proceso. Buscando asuntos que contengan: '{asuntoBusqueda}'");

            Directory.CreateDirectory(config.FolderTemporal);
            foreach (var f in Directory.GetFiles(config.FolderTemporal))
                try { File.Delete(f); } catch { }
            Log("Carpeta temporal limpiada.");

            string? rutaAdjunto = null;
            await Task.Run(() =>
            {
                using var outlook = new OutlookService();
                rutaAdjunto = outlook.BuscarYGuardarAdjunto(
                    config.OutlookCarpeta, asuntoBusqueda, config.FolderTemporal,
                    string.IsNullOrWhiteSpace(config.OutlookCuenta) ? null : config.OutlookCuenta.Trim(),
                    Log);
            }, _cts.Token);

            if (string.IsNullOrEmpty(rutaAdjunto))
            {
                Log("No se encontró correo. Enviando aviso al destinatario.");
                var cuerpoAviso =
                    $"<p>No se ha encontrado correo cuyo asunto contenga: " +
                    $"'{System.Net.WebUtility.HtmlEncode(asuntoBusqueda)}', " +
                    $"se procede a detener el proceso.</p>" +
                    $"<p>Saludos cordiales.</p>" +
                    $"<p><strong> - Notificación automática. - </strong></p>";
                using var outlook = new OutlookService();
                outlook.EnviarCorreo(config.CorreoTo,
                    "ROBOT | No se encontró correo de bajas", cuerpoAviso, Log);
                return;
            }

            var rutaDestinoUser = Path.Combine(
                config.FolderUser.TrimEnd('\\', '/'), $"{DateTime.Now:dd.MM.yy}.xlsx");
            Directory.CreateDirectory(config.FolderUser);
            if (File.Exists(rutaDestinoUser)) File.Delete(rutaDestinoUser);
            File.Move(rutaAdjunto, rutaDestinoUser);
            Log($"Archivo movido a: {rutaDestinoUser}");

            string? htmlTabla = null;
            await Task.Run(() =>
            {
                var svc = new ProcessService();
                htmlTabla = svc.ProcesarBajas(rutaDestinoUser, config, Log);
            }, _cts.Token);

            if (!string.IsNullOrEmpty(config.CorreoTo))
            {
                using var outlook = new OutlookService();
                if (!string.IsNullOrEmpty(htmlTabla))
                {
                    var asunto = $"{config.AsuntoCorreoS} - {DateTime.Now:dd/MM/yyyy}";
                    outlook.EnviarCorreo(config.CorreoTo, asunto, htmlTabla, Log);
                }
                else
                {
                    var cuerpoSinBajas =
                        "<p>Buenas tardes estimados,</p><p></p>" +
                        "<p>En el procesamiento del día <strong>" + DateTime.Now.ToString("dd/MM/yyyy") + "</strong> " +
                        "no se encontraron cuentas a dar de baja.</p><p></p>" +
                        "<p>Saludos cordiales.</p><p></p>" +
                        "<p><strong> - Esta es una notificación automática, por favor, no responder este correo. - </strong></p>";
                    outlook.EnviarCorreo(config.CorreoTo,
                        "ROBOT | No se encontraron bajas - " + DateTime.Now.ToString("dd/MM/yyyy"),
                        cuerpoSinBajas, Log);
                }
            }

            Log("Proceso finalizado correctamente.");
        }
        catch (OperationCanceledException)
        {
            Log("Proceso cancelado por el usuario.");
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
            try
            {
                var config = _configPanel.SaveTo();
                var cuerpoError =
                    $"<p>Se ha presentado el siguiente error: " +
                    $"{System.Net.WebUtility.HtmlEncode(ex.ToString())}. " +
                    $"Se procede a detener el proceso.</p>" +
                    $"<p>Saludos cordiales.</p>" +
                    $"<p><strong> - Notificación automática. - </strong></p>";
                using var outlook = new OutlookService();
                outlook.EnviarCorreo(config.CorreoTo,
                    "ROBOT | ERROR | Notificador Bajas", cuerpoError, Log);
            }
            catch { }
        }
        finally
        {
            _btnRun.Enabled    = true;
            _btnCancel.Visible = false;
            _cts?.Dispose();
        }
    }

    private void AsignarIcono()
    {
        try
        {
            var baseDir = Application.StartupPath;
            var iconDir = Path.Combine(baseDir, "icon");
            if (!Directory.Exists(iconDir)) return;

            // Preferir app.ico, luego app.png, luego cualquier .ico o .png en icon\
            var icoPath = Path.Combine(iconDir, "app.ico");
            if (File.Exists(icoPath))
            {
                using var ico = new Icon(icoPath);
                Icon = (Icon)ico.Clone();
                return;
            }

            var pngPath = Path.Combine(iconDir, "app.png");
            if (File.Exists(pngPath))
            {
                UsarIconoDesdePng(pngPath);
                return;
            }

            var primerIco = Directory.GetFiles(iconDir, "*.ico").FirstOrDefault();
            if (primerIco != null)
            {
                using var ico = new Icon(primerIco);
                Icon = (Icon)ico.Clone();
                return;
            }

            var primerPng = Directory.GetFiles(iconDir, "*.png").FirstOrDefault();
            if (primerPng != null)
                UsarIconoDesdePng(primerPng);
        }
        catch { /* icono opcional */ }
    }

    private void UsarIconoDesdePng(string pngPath)
    {
        // La barra de tareas usa iconos 32x32 (y 16x16). Un PNG grande no se muestra bien;
        // redimensionamos a 32x32 para que el botón de la barra de tareas lo muestre correctamente.
        using var bmpOrig = new Bitmap(pngPath);
        using var bmp = new Bitmap(bmpOrig, new Size(32, 32));
        using var icon = Icon.FromHandle(bmp.GetHicon());
        Icon = (Icon)icon.Clone();
    }
}
