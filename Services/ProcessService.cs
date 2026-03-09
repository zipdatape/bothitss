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

    /// <summary>Indica si el archivo parece un ZIP/Excel por la firma PK (0x50 0x4B).</summary>
    private static bool EsArchivoZip(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (fs.Length < 2) return false;
            int b0 = fs.ReadByte();
            int b1 = fs.ReadByte();
            return b0 == 0x50 && b1 == 0x4B; // PK
        }
        catch { return false; }
    }

    /// <summary>Obtiene los DNI de bajas desde el adjunto: Excel (.xlsx) o CSV (misma columna DNI, índice 1).</summary>
    private static HashSet<string> ObtenerDnisBajaDesdeArchivo(string rutaArchivo, Action<string>? log, int colDni = 1, string? sheetName = null)
    {
        var dnisBaja = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (EsArchivoZip(rutaArchivo))
        {
            using var book = new XLWorkbook(rutaArchivo);
            var sheet = !string.IsNullOrWhiteSpace(sheetName) ? book.Worksheet(sheetName) : null;
            sheet ??= book.Worksheets.FirstOrDefault();
            if (sheet == null) return dnisBaja;
            var rows = sheet.RangeUsed()?.Rows().Skip(1).ToList() ?? new List<IXLRangeRow>();
            foreach (var row in rows)
            {
                var dni = row.Cell(colDni + 1).GetString().Trim();
                var dniNorm = NormalizarDni(dni);
                if (!string.IsNullOrEmpty(dniNorm)) dnisBaja.Add(dniNorm);
            }
            var sampleExcel = string.Join(", ", dnisBaja.Take(10));
            log?.Invoke($"Leído como Excel. Filas de bajas: {dnisBaja.Count}. DNIs (muestra): {sampleExcel}");
            return dnisBaja;
        }
        // No es ZIP: intentar como CSV (p. ej. adjunto con extensión .xlsx pero contenido CSV)
        try
        {
            var enc = Encoding.GetEncoding("iso-8859-15");
            var primeraLinea = File.ReadLines(rutaArchivo, enc).FirstOrDefault() ?? "";
            // Detectar delimitador: priorizar el que genere más columnas (tab, ';' o ',')
            int tabs   = primeraLinea.Split('\t').Length;
            int semis  = primeraLinea.Split(';').Length;
            int comas  = primeraLinea.Split(',').Length;
            string delimiter;
            if (tabs >= semis && tabs >= comas && primeraLinea.Contains('\t'))
                delimiter = "\t";
            else if (semis >= tabs && semis >= comas && primeraLinea.Contains(';'))
                delimiter = ";";
            else
                delimiter = ",";
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true, Delimiter = delimiter };
            using var reader = new StreamReader(rutaArchivo, enc);
            using var csv = new CsvReader(reader, csvConfig);
            csv.Read();
            csv.ReadHeader();
            while (csv.Read())
            {
                if (csv.TryGetField(colDni, out string? dni))
                {
                    var dniNorm = NormalizarDni(dni);
                    if (!string.IsNullOrEmpty(dniNorm)) dnisBaja.Add(dniNorm);
                }
            }
            var delimName = delimiter == "\t" ? "tab" : delimiter == ";" ? "punto y coma" : "coma";
            var sampleCsv = string.Join(", ", dnisBaja.Take(10));
            log?.Invoke($"Leído como CSV (delimitador: {delimName}). Filas de bajas: {dnisBaja.Count}. DNIs (muestra): {sampleCsv}");
            return dnisBaja;
        }
        catch (Exception ex)
        {
            log?.Invoke($"No se pudo leer como CSV: {ex.Message}");
            throw;
        }
    }

    /// <summary>Procesa el Excel/CSV de bajas (ya en FolderUser con nombre por fecha), actualiza la base CSV y devuelve el HTML de la tabla para el correo.</summary>
    public string? ProcesarBajas(string rutaExcel, AppConfig config, Action<string>? log = null)
    {
        try
        {
            if (!File.Exists(rutaExcel))
            {
                log?.Invoke($"No existe el archivo: {rutaExcel}");
                return null;
            }

            // DNI en columna índice 1 del archivo de bajas (plantilla Excel/CSV de bajas: ID SAP, DNI, ...).
            var dnisBaja = ObtenerDnisBajaDesdeArchivo(rutaExcel, log, colDni: 1, sheetName: config.SheetName);

            var pathBase = Path.Combine(config.FolderBASE.TrimEnd('\\', '/'), config.FileBase);
            if (!File.Exists(pathBase))
            {
                log?.Invoke($"No existe el archivo base: {pathBase}");
                return null;
            }

            // La BASE HITSS.csv está separada por coma (,), sin cabecera, DNI en columna 0.
            // Se procesa en streaming para evitar cargar todo el archivo en RAM.
            var enc = Encoding.GetEncoding(1252);
            var csvCfg = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = false,
                Delimiter = ",",
                BadDataFound = null,
                MissingFieldFound = null,
                ShouldQuote = _ => false,
            };

            // Columna 0 = DNI en la BASE HITSS.csv (sin cabecera, delimitada por coma).
            const int colDniBase = 0;
            var bajas = new List<string[]>();
            int totalFilas = 0;
            int coincidencias = 0;

            // Backup antes de modificar
            var folderBkp = config.FolderBCKP.TrimEnd('\\', '/');
            Directory.CreateDirectory(folderBkp);
            var pathBkp = Path.Combine(folderBkp, $"{config.FileBkp}{DateTime.Now:ddMMyyyy}.csv");
            File.Copy(pathBase, pathBkp, true);
            log?.Invoke($"Backup guardado: {pathBkp}");

            // Streaming: leer original y escribir filtrado a archivo temporal en paralelo.
            var pathTemp = pathBase + ".tmp";
            try
            {
                string? primerDniEjemplo = null;
                using (var reader = new StreamReader(pathBase, enc))
                using (var csvR   = new CsvReader(reader, csvCfg))
                using (var writer = new StreamWriter(pathTemp, false, enc))
                using (var csvW   = new CsvWriter(writer, csvCfg))
                {
                    while (csvR.Read())
                    {
                        var row = new List<string>();
                        for (int i = 0; csvR.TryGetField(i, out string? v); i++)
                            row.Add(v ?? "");

                        totalFilas++;
                        var arr    = row.ToArray();
                        var dniRaw = arr.Length > colDniBase ? arr[colDniBase] : "";
                        var dni    = NormalizarDni(dniRaw);
                        if (primerDniEjemplo == null) primerDniEjemplo = dniRaw;

                        if (!string.IsNullOrEmpty(dni) && dnisBaja.Contains(dni))
                        {
                            bajas.Add(arr);
                            coincidencias++;
                        }
                        else
                        {
                            foreach (var cell in arr) csvW.WriteField(cell);
                            csvW.NextRecord();
                        }
                    }
                }

                log?.Invoke($"Base HITSS leída. Filas totales: {totalFilas}. Ejemplo DNI primera fila: {primerDniEjemplo ?? "<sin datos>"}");
                log?.Invoke($"Cruce con base: DNIs en bajas = {dnisBaja.Count}, filas en base = {totalFilas}, coincidencias = {coincidencias}.");

                // Reemplazar original con el filtrado
                File.Delete(pathBase);
                File.Move(pathTemp, pathBase);
            }
            catch
            {
                // Si algo falló, eliminar el temporal para no dejar archivos huérfanos
                if (File.Exists(pathTemp)) try { File.Delete(pathTemp); } catch { }
                throw;
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
