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
            await DownloadAndReplaceAsync(release.ExeUrl, release.PageUrl, log);
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

    /// <summary>Si existe el marcador de "actualización en curso", la reemplazo falló. Devuelve la URL de la release para abrir.</summary>
    public static string? ConsumeFailedUpdateMarker()
    {
        try
        {
            var dir = Path.GetDirectoryName(Environment.ProcessPath!);
            if (string.IsNullOrEmpty(dir)) return null;
            var marker = Path.Combine(dir, ".updating_to");
            if (!File.Exists(marker)) return null;
            var url = File.ReadAllText(marker).Trim();
            File.Delete(marker);
            return string.IsNullOrEmpty(url) ? null : url;
        }
        catch { return null; }
    }

    private static async Task DownloadAndReplaceAsync(string exeUrl, string pageUrl, Action<string> log)
    {
        try
        {
            var currentExe = Environment.ProcessPath!;
            var dir        = Path.GetDirectoryName(currentExe)!;
            var newExe     = Path.Combine(dir, "NotificadorBajasHitss_update.exe");

            log("Descargando actualización...");
            var data = await _http.GetByteArrayAsync(exeUrl);
            await File.WriteAllBytesAsync(newExe, data);

            // Marcador: si el reemplazo falla, al reiniciar no volver a ofrecer la actualización en bucle
            var markerPath = Path.Combine(dir, ".updating_to");
            await File.WriteAllTextAsync(markerPath, pageUrl);

            // Bat: esperar a que el proceso termine, reemplazar; solo reiniciar si move tuvo éxito.
            // Si move falla (ej. sin permisos), no reiniciamos el exe antiguo → se evita el bucle.
            var batPath = Path.Combine(Path.GetTempPath(), "bothitss_updater.bat");
            var pid     = Environment.ProcessId;
            await File.WriteAllTextAsync(batPath,
                "@echo off\r\n" +
                "timeout /t 2 /nobreak >nul\r\n" +
                ":wait\r\n" +
                $"tasklist /fi \"PID eq {pid}\" 2>nul | find \"{pid}\" >nul\r\n" +
                "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto wait)\r\n" +
                $"move /y \"{newExe}\" \"{currentExe}\"\r\n" +
                "if errorlevel 1 (\r\n" +
                "  start \"\" \"" + pageUrl.Replace("\"", "") + "\"\r\n" +
                ") else (\r\n" +
                "  del \"" + markerPath + "\" 2>nul\r\n" +
                "  start \"\" \"" + currentExe + "\"\r\n" +
                ")\r\n" +
                "del \"%~f0\"\r\n");

            log("Aplicando actualización. La aplicación se cerrará...");

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
