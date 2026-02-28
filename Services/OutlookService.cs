using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace NotificadorBajasHitssApp.Services;

public class OutlookService : IDisposable
{
    private Microsoft.Office.Interop.Outlook.Application? _app;

    public void Log(string message, Action<string>? log = null) => log?.Invoke(message);

    /// <summary>Busca el primer correo no leído, con adjunto, cuyo asunto contenga asuntoBusqueda.</summary>
    /// <returns>Ruta del archivo adjunto guardado, o null si no se encontró.</returns>
    public string? BuscarYGuardarAdjunto(string carpetaOutlook, string asuntoBusqueda, string carpetaDestino, Action<string>? log = null)
    {
        try
        {
            _app = new Microsoft.Office.Interop.Outlook.Application();
            var ns = _app.GetNamespace("MAPI");
            MAPIFolder? folder = null;
            try
            {
                folder = GetFolderByPath(ns, carpetaOutlook);
            }
            catch (System.Exception ex)
            {
                Log($"No se pudo abrir la carpeta de Outlook: {carpetaOutlook}. {ex.Message}", log);
                return null;
            }

            if (folder == null)
            {
                Log($"Carpeta no encontrada: {carpetaOutlook}", log);
                return null;
            }

            Directory.CreateDirectory(carpetaDestino);
            var items = folder.Items;
            items.Sort("[ReceivedTime]", true);

            for (int i = items.Count; i >= 1; i--)
            {
                try
                {
                    var item = items[i];
                    if (item is not MailItem mail) continue;
                    if (mail.UnRead != true) continue;
                    if (!mail.Subject.Contains(asuntoBusqueda, StringComparison.OrdinalIgnoreCase)) continue;
                    if (mail.Attachments.Count == 0) continue;

                    // Guardar primer adjunto
                    var adj = mail.Attachments[1];
                    var ext = Path.GetExtension(adj.FileName);
                    if (string.IsNullOrEmpty(ext)) ext = ".xlsx";
                    var destPath = Path.Combine(carpetaDestino, Path.GetFileNameWithoutExtension(adj.FileName) + ext);
                    adj.SaveAsFile(destPath);
                    mail.UnRead = false;
                    mail.Save();
                    Marshal.ReleaseComObject(mail);
                    Log($"Correo encontrado. Adjunto guardado: {destPath}", log);
                    return destPath;
                }
                catch (System.Exception ex)
                {
                    Log($"Error al procesar correo: {ex.Message}", log);
                }
            }

            Log("No se encontró ningún correo que cumpla los criterios (no leído, con adjunto, asunto con la fecha).", log);
            return null;
        }
        catch (System.Exception ex)
        {
            Log($"Error al conectar con Outlook: {ex.Message}", log);
            return null;
        }
        finally
        {
            if (_app != null)
            {
                Marshal.ReleaseComObject(_app);
                _app = null;
            }
        }
    }

    /// <summary>Envía un correo con el cliente por defecto de Outlook.</summary>
    public bool EnviarCorreo(string para, string asunto, string cuerpoHtml, Action<string>? log = null)
    {
        try
        {
            _app = new Microsoft.Office.Interop.Outlook.Application();
            var mail = _app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as MailItem;
            if (mail == null) { Log("No se pudo crear el mensaje.", log); return false; }
            mail.To = para;
            mail.Subject = asunto;
            mail.HTMLBody = cuerpoHtml;
            mail.Send();
            Marshal.ReleaseComObject(mail);
            Log($"Correo enviado a {para}.", log);
            return true;
        }
        catch (System.Exception ex)
        {
            Log($"Error al enviar correo: {ex.Message}", log);
            return false;
        }
        finally
        {
            if (_app != null)
            {
                Marshal.ReleaseComObject(_app);
                _app = null;
            }
        }
    }

    private static MAPIFolder? GetFolderByPath(NameSpace ns, string path)
    {
        var parts = path.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
        MAPIFolder? folder = ns.Folders[1];
        for (int i = 0; i < parts.Length; i++)
        {
            var name = parts[i];
            MAPIFolder? found = null;
            foreach (MAPIFolder f in folder.Folders)
            {
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
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

    public void Dispose() => _app = null;
}
