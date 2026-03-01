using System.Runtime.InteropServices;

namespace NotificadorBajasHitssApp.Services;

/// <summary>
/// Servicio de Outlook usando COM late-binding (dynamic).
/// No requiere Microsoft.Office.Interop.Outlook ni office.dll en el GAC.
/// Compatible con todas las versiones de Outlook (2013/2016/2019/365 MSI y Click-to-Run).
/// </summary>
public class OutlookService : IDisposable
{
    // OlItemType.olMailItem = 0  (para CreateItem)
    // OlObjectClass.olMail  = 43 (para verificar tipo de ítem)
    private const int OlMailItem  = 0;
    private const int OlMailClass = 43;

    private dynamic? _app;

    // ── Detección de Outlook y cuentas ───────────────────────────────────────

    /// <summary>Devuelve true si Outlook está instalado y registrado como COM.</summary>
    public static bool OutlookDisponible()
        => Type.GetTypeFromProgID("Outlook.Application") != null;

    /// <summary>Devuelve las cuentas SMTP configuradas en el perfil de Outlook activo.</summary>
    public static List<string> ObtenerCuentas()
    {
        var lista = new List<string>();
        try
        {
            var t = Type.GetTypeFromProgID("Outlook.Application");
            if (t == null) return lista;

            dynamic app = Activator.CreateInstance(t)!;
            var ns       = app.GetNamespace("MAPI");
            var accounts = ns.Accounts;
            int total    = (int)accounts.Count;

            for (int i = 1; i <= total; i++)
            {
                var acc   = accounts[i];
                var smtp  = (string)acc.SmtpAddress;
                var name  = (string)acc.DisplayName;
                lista.Add(string.IsNullOrWhiteSpace(smtp) ? name : $"{name} <{smtp}>");
            }
            Marshal.ReleaseComObject(app);
        }
        catch { /* Outlook no disponible o sin perfil */ }
        return lista;
    }

    /// <summary>Devuelve los nombres de las carpetas raíz (cuentas/almacenes) de Outlook para elegir en qué cuenta buscar.</summary>
    public static List<string> ObtenerCarpetasRaiz()
    {
        var lista = new List<string>();
        try
        {
            var t = Type.GetTypeFromProgID("Outlook.Application");
            if (t == null) return lista;

            dynamic app = Activator.CreateInstance(t)!;
            var ns     = app.GetNamespace("MAPI");
            var roots  = ns.Folders;
            int total  = (int)roots.Count;

            for (int i = 1; i <= total; i++)
            {
                var folder = roots[i];
                var name   = (string)folder.Name;
                if (!string.IsNullOrWhiteSpace(name))
                    lista.Add(name.Trim());
            }
            Marshal.ReleaseComObject(app);
        }
        catch { /* Outlook no disponible o sin perfil */ }
        return lista;
    }

    // ── Búsqueda de correo y guardado del adjunto ─────────────────────────────

    public string? BuscarYGuardarAdjunto(
        string carpetaOutlook,
        string asuntoBusqueda,
        string carpetaDestino,
        string? cuentaOutlook = null,
        Action<string>? log = null)
    {
        try
        {
            var t = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Outlook no está instalado o no está registrado como COM.");

            _app = Activator.CreateInstance(t)!;
            var ns = _app.GetNamespace("MAPI");

            dynamic? folder;
            try   { folder = GetFolderByPath(ns, carpetaOutlook, cuentaOutlook); }
            catch (Exception ex)
            {
                log?.Invoke($"No se pudo abrir la carpeta de Outlook '{carpetaOutlook}': {ex.Message}");
                return null;
            }

            if (folder == null)
            {
                log?.Invoke($"Carpeta no encontrada en Outlook: {carpetaOutlook}");
                return null;
            }

            Directory.CreateDirectory(carpetaDestino);

            var items = folder.Items;
            items.Sort("[ReceivedTime]", true);
            int total = (int)items.Count;

            log?.Invoke($"Revisando {total} correo(s) en la carpeta '{carpetaOutlook}'...");

            for (int i = total; i >= 1; i--)
            {
                try
                {
                    var item = items[i];
                    if ((int)item.Class != OlMailClass)       continue;
                    if (!(bool)item.UnRead)                   continue;
                    if ((int)item.Attachments.Count == 0)     continue;

                    var subject = (string)item.Subject;
                    if (!subject.Contains(asuntoBusqueda, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var adj      = item.Attachments[1];
                    var fileName = (string)adj.FileName;
                    var ext      = Path.GetExtension(fileName);
                    if (string.IsNullOrEmpty(ext)) ext = ".xlsx";

                    var destPath = Path.Combine(carpetaDestino,
                        Path.GetFileNameWithoutExtension(fileName) + ext);
                    adj.SaveAsFile(destPath);
                    item.UnRead = false;
                    item.Save();
                    Marshal.ReleaseComObject(item);

                    log?.Invoke($"Correo encontrado: '{subject}'. Adjunto guardado: {destPath}");
                    return destPath;
                }
                catch (Exception ex)
                {
                    log?.Invoke($"Error al procesar correo #{i}: {ex.Message}");
                }
            }

            log?.Invoke("No se encontró correo no leído con adjunto y asunto coincidente.");
            return null;
        }
        catch (Exception ex)
        {
            log?.Invoke($"Error al conectar con Outlook: {ex.Message}");
            return null;
        }
        finally { LiberarApp(); }
    }

    // ── Envío de correo ───────────────────────────────────────────────────────

    public bool EnviarCorreo(
        string para,
        string asunto,
        string cuerpoHtml,
        Action<string>? log = null)
    {
        try
        {
            var t = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Outlook no está instalado.");

            _app = Activator.CreateInstance(t)!;
            var mail = _app.CreateItem(OlMailItem);
            mail.To       = para;
            mail.Subject  = asunto;
            mail.HTMLBody = cuerpoHtml;
            mail.Send();
            Marshal.ReleaseComObject(mail);
            log?.Invoke($"Correo enviado a: {para}");
            return true;
        }
        catch (Exception ex)
        {
            log?.Invoke($"Error al enviar correo: {ex.Message}");
            return false;
        }
        finally { LiberarApp(); }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static dynamic? GetFolderByPath(dynamic ns, string path, string? cuentaRaiz = null)
    {
        var parts = path.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
        dynamic folder;

        if (string.IsNullOrWhiteSpace(cuentaRaiz))
        {
            folder = ns.Folders[1];
        }
        else
        {
            var roots = ns.Folders;
            int total = (int)roots.Count;
            dynamic? found = null;
            for (int i = 1; i <= total; i++)
            {
                var f = roots[i];
                if (string.Equals((string)f.Name, cuentaRaiz.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    found = f;
                    break;
                }
            }
            if (found == null)
                return null;
            folder = found;
        }

        foreach (var part in parts)
        {
            dynamic? found      = null;
            var      subFolders = folder.Folders;
            int      count      = (int)subFolders.Count;

            for (int j = 1; j <= count; j++)
            {
                dynamic f = subFolders[j];
                if (string.Equals((string)f.Name, part, StringComparison.OrdinalIgnoreCase))
                {
                    found = f;
                    break;
                }
            }
            if (found == null) return null;
            folder = found;
        }
        return folder;
    }

    private void LiberarApp()
    {
        if (_app == null) return;
        try { Marshal.ReleaseComObject(_app); } catch { }
        _app = null;
    }

    public void Dispose() => LiberarApp();
}
