using System.Reflection;
using System.Text.Json;

namespace NotificadorBajasHitssApp.Services;

public static class UpdateService
{
    private const string Owner = "zipdatape";
    private const string Repo  = "bothitss";

    private static readonly HttpClient _http = new()
    {
        Timeout = TimeSpan.FromSeconds(12)
    };

    static UpdateService()
    {
        _http.DefaultRequestHeaders.Add("User-Agent", "NotificadorBajasHitss-Updater");
    }

    public static Version CurrentVersion =>
        Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0);

    public record ReleaseInfo(bool HasUpdate, string Tag, string ExeUrl, string PageUrl);

    /// <summary>Consulta la última release en GitHub. Devuelve info de actualización si hay versión más nueva.</summary>
    public static async Task<ReleaseInfo> CheckAsync()
    {
        try
        {
            var json = await _http.GetStringAsync(
                $"https://api.github.com/repos/{Owner}/{Repo}/releases/latest");

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            var tag     = root.GetProperty("tag_name").GetString() ?? "";
            var pageUrl = root.GetProperty("html_url").GetString() ?? "";
            var verStr  = tag.TrimStart('v', 'V');

            if (!Version.TryParse(verStr, out var remote) || remote <= CurrentVersion)
                return new ReleaseInfo(false, tag, "", pageUrl);

            // Buscar asset .exe en la release
            var exeUrl = "";
            if (root.TryGetProperty("assets", out var assets))
            {
                foreach (var asset in assets.EnumerateArray())
                {
                    var name = asset.GetProperty("name").GetString() ?? "";
                    if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        exeUrl = asset.GetProperty("browser_download_url").GetString() ?? "";
                        break;
                    }
                }
            }

            return new ReleaseInfo(true, tag, exeUrl, pageUrl);
        }
        catch
        {
            // Sin conexión o sin releases → ignorar silenciosamente
            return new ReleaseInfo(false, "", "", "");
        }
    }

    /// <summary>
    /// Descarga el .exe de la nueva versión y lo aplica mediante un script de actualización.
    /// Si no hay asset directo, abre el navegador en la página de la release.
    /// </summary>
    public static async Task ApplyAsync(ReleaseInfo release, Action<string> log)
    {
        if (!string.IsNullOrEmpty(release.ExeUrl))
        {
            await DownloadAndReplaceAsync(release.ExeUrl, log);
        }
        else
        {
            // Sin asset .exe → abrir página de descarga
            log("Abriendo página de descarga en el navegador...");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName        = release.PageUrl,
                UseShellExecute = true
            });
        }
    }

    private static async Task DownloadAndReplaceAsync(string exeUrl, Action<string> log)
    {
        try
        {
            var currentExe = Environment.ProcessPath!;
            var dir        = Path.GetDirectoryName(currentExe)!;
            var newExe     = Path.Combine(dir, "NotificadorBajasHitss_update.exe");

            log("Descargando actualización...");
            var data = await _http.GetByteArrayAsync(exeUrl);
            await File.WriteAllBytesAsync(newExe, data);

            // Script bat que espera a que el proceso termine, reemplaza el exe y reinicia
            var batPath = Path.Combine(Path.GetTempPath(), "bothitss_updater.bat");
            var pid     = Environment.ProcessId;

            await File.WriteAllTextAsync(batPath,
                $"@echo off\r\n" +
                $":wait\r\n" +
                $"tasklist /fi \"PID eq {pid}\" 2>nul | find \"{pid}\" >nul\r\n" +
                $"if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\r\n" +
                $"move /y \"{newExe}\" \"{currentExe}\"\r\n" +
                $"start \"\" \"{currentExe}\"\r\n" +
                $"del \"%~f0\"\r\n");

            log("Aplicando actualización. La aplicación se reiniciará...");

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName       = "cmd.exe",
                Arguments      = $"/c \"{batPath}\"",
                WindowStyle    = System.Diagnostics.ProcessWindowStyle.Hidden,
                CreateNoWindow = true
            });

            Application.Exit();
        }
        catch (Exception ex)
        {
            log($"Error al aplicar actualización: {ex.Message}");
        }
    }
}
