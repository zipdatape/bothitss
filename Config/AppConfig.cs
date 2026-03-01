namespace NotificadorBajasHitssApp.Config;

/// <summary>
/// Configuración del proceso (equivalente a Config.xlsx).
/// Se guarda/carga desde config.json en la carpeta de la aplicación.
/// </summary>
public class AppConfig
{
    public string Proceso { get; set; } = "Notificador de Bajas de Usuarios Hitss";
    public int FechaDia { get; set; } = 1;
    public string FolderTemporal { get; set; } = "";
    public string FolderUser { get; set; } = "";
    /// <summary>Texto fijo del asunto a buscar (ej. "CESE DE PERSONAL - "). La fecha del correo es dinámica.</summary>
    public string AsuntoCorreoR { get; set; } = "CESE DE PERSONAL - ";
    public string SheetName { get; set; } = "Hoja1";
    public string FolderBASE { get; set; } = "";
    public string FileBase { get; set; } = "BASE HITSS.csv";
    public string FolderBCKP { get; set; } = "";
    public string FileBkp { get; set; } = "BASE HITSS BKP";
    public string CorreoTo { get; set; } = "";
    public string AsuntoCorreoS { get; set; } = "Notificación de Bajas";
    /// <summary>Cuenta/carpeta raíz de Outlook donde buscar (vacío = primera cuenta).</summary>
    public string OutlookCuenta { get; set; } = "";
    /// <summary>Carpeta de Outlook donde buscar (ej. "Bandeja de entrada\\C.H_BAJAS").</summary>
    public string OutlookCarpeta { get; set; } = "Bandeja de entrada\\C.H_BAJAS";

    public Dictionary<string, string> ToDictionary()
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["xProceso"] = Proceso,
            ["xFechaDia"] = FechaDia.ToString(),
            ["xFolderTemporal"] = EnsureTrailingSlash(FolderTemporal),
            ["xFolderUser"] = EnsureTrailingSlash(FolderUser),
            ["xAsuntoCorreoR"] = AsuntoCorreoR,
            ["xSheetName"] = SheetName,
            ["xFolderBASE"] = EnsureTrailingSlash(FolderBASE),
            ["xFileBase"] = FileBase,
            ["xFolderBCKP"] = EnsureTrailingSlash(FolderBCKP),
            ["xFileBkp"] = FileBkp,
            ["xCorreoTo"] = CorreoTo,
            ["xAsuntoCorreoS"] = AsuntoCorreoS,
        };
    }

    private static string EnsureTrailingSlash(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return path;
        return path.TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
    }
}
