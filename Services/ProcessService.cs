using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using NotificadorBajasHitssApp.Config;

namespace NotificadorBajasHitssApp.Services;

public class ProcessService
{
    /// <summary>Normaliza DNI: si cumple ^0[0-9]*$ quita "00" o "0" al inicio.</summary>
    private static string NormalizarDni(string? dni)
    {
        if (string.IsNullOrWhiteSpace(dni)) return dni ?? "";
        var s = dni.Trim();
        if (System.Text.RegularExpressions.Regex.IsMatch(s, @"^0[0-9]*$"))
        {
            if (s.StartsWith("00") && s.Length > 2) return s.Substring(2);
            if (s.StartsWith("0") && s.Length > 1) return s.Substring(1);
        }
        return s;
    }

    /// <summary>Procesa el Excel de bajas (ya en FolderUser con nombre por fecha), actualiza la base CSV y devuelve el HTML de la tabla para el correo.</summary>
    public string? ProcesarBajas(string rutaExcel, AppConfig config, Action<string>? log = null)
    {
        try
        {
            if (!File.Exists(rutaExcel))
            {
                log?.Invoke($"No existe el archivo: {rutaExcel}");
                return null;
            }

            using var book = new XLWorkbook(rutaExcel);
            var sheet = book.Worksheet(config.SheetName) ?? book.Worksheets.First();
            var rows = sheet.RangeUsed()?.Rows().Skip(1).ToList() ?? new List<IXLRangeRow>(); // Skip header

            var pathBase = Path.Combine(config.FolderBASE.TrimEnd('\\', '/'), config.FileBase);
            if (!File.Exists(pathBase))
            {
                log?.Invoke($"No existe el archivo base: {pathBase}");
                return null;
            }

            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = false };
            var baseRows = new List<string[]>();
            using (var reader = new StreamReader(pathBase, Encoding.GetEncoding("iso-8859-15")))
            using (var csv = new CsvReader(reader, csvConfig))
            {
                while (csv.Read())
                {
                    var row = new List<string>();
                    for (int i = 0; csv.TryGetField(i, out string? v); i++)
                        row.Add(v ?? "");
                    baseRows.Add(row.ToArray());
                }
            }

            int colDniExcel = 1; // Columna del DNI en el Excel (0-based)
            int colDniBase = 1;  // Columna del DNI en el CSV base (0-based)
            var bajas = new List<string[]>();
            var dnisBaja = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                var dniExcel = row.Cell(colDniExcel + 1).GetString().Trim(); // ClosedXML 1-based
                var dniNorm = NormalizarDni(dniExcel);
                if (string.IsNullOrEmpty(dniNorm)) continue;
                dnisBaja.Add(dniNorm);
            }

            var nuevaBase = new List<string[]>();
            foreach (var row in baseRows)
            {
                if (row.Length <= colDniBase) continue;
                var dniBase = NormalizarDni(row[colDniBase]);
                if (dnisBaja.Contains(dniBase))
                    bajas.Add(row);
                else
                    nuevaBase.Add(row);
            }

            // Backup
            var folderBkp = config.FolderBCKP.TrimEnd('\\', '/');
            Directory.CreateDirectory(folderBkp);
            var pathBkp = Path.Combine(folderBkp, $"{config.FileBkp}{DateTime.Now:ddMMyyyy}.csv");
            File.Copy(pathBase, pathBkp, true);
            log?.Invoke($"Backup guardado: {pathBkp}");

            // Escribir nueva base
            using (var writer = new StreamWriter(pathBase, false, Encoding.GetEncoding(1252)))
            using (var csv = new CsvWriter(writer, csvConfig))
            {
                foreach (var row in nuevaBase)
                {
                    foreach (var cell in row)
                        csv.WriteField(cell);
                    csv.NextRecord();
                }
            }
            log?.Invoke($"Base actualizada. Bajas encontradas: {bajas.Count}");

            if (bajas.Count == 0)
                return null;

            return BuildTablaHtmlBajas(bajas);
        }
        catch (Exception ex)
        {
            log?.Invoke($"Error en ProcesarBajas: {ex.Message}");
            throw;
        }
    }

    private static string BuildTablaHtmlBajas(List<string[]> bajas)
    {
        var sb = new StringBuilder();
        sb.Append("<p>Buenas tardes estimados,</p><p></p>");
        sb.Append("<p>Se solicita de su apoyo para proceder con la baja de las siguientes cuentas E,</p><p></p>");
        sb.Append("<table class='demoTable' border='0' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border: 1px solid black;' rules='all'><tbody>");
        foreach (var row in bajas)
        {
            sb.Append("<tr>");
            for (int i = 0; i < Math.Min(6, row.Length); i++)
                sb.Append($"<td style='border: 2px solid #fd9550;padding:0cm 5.4pt;height:15.0pt;'>{System.Net.WebUtility.HtmlEncode(row[i])}</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table><p></p><p>Saludos cordiales.</p><p></p>");
        sb.Append("<p><strong> - Esta es una notificación automática, por favor, no responder este correo. - </strong></p>");
        return sb.ToString();
    }
}
