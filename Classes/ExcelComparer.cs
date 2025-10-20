using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace FocaExcelExport.Classes
{
    public sealed class CompareProgress
    {
        public int CurrentStep { get; set; }
        public int TotalSteps { get; set; }
        public string Message { get; set; }
    }

    public class ExcelComparer
    {
        public async Task CompareAsync(string basePath, string newPath, string outputPath, Action<CompareProgress> report, string keyMode = "URL", bool highlight = true)
        {
            await Task.Run(() => CompareInternal(basePath, newPath, outputPath, report, keyMode, highlight));
        }

        private void CompareInternal(string basePath, string newPath, string outputPath, Action<CompareProgress> report, string keyMode, bool highlight)
        {
            report?.Invoke(new CompareProgress { CurrentStep = 0, TotalSteps = 100, Message = "Leyendo libros..." });

            using (var baseWb = new XLWorkbook(basePath))
            using (var newWb = new XLWorkbook(newPath))
            using (var outWb = new XLWorkbook())
            {
                // Colecciones de hojas
                var baseSheets = baseWb.Worksheets.ToDictionary(ws => ws.Name, StringComparer.OrdinalIgnoreCase);
                var newSheets = newWb.Worksheets.ToDictionary(ws => ws.Name, StringComparer.OrdinalIgnoreCase);

                // Conjunto total de hojas
                var allSheetNames = new HashSet<string>(baseSheets.Keys, StringComparer.OrdinalIgnoreCase);
                foreach (var n in newSheets.Keys) allSheetNames.Add(n);

                int processed = 0;
                int total = allSheetNames.Count;

                foreach (var sheetName in allSheetNames)
                {
                    processed++;
                    report?.Invoke(new CompareProgress
                    {
                        CurrentStep = processed,
                        TotalSteps = total + 1,
                        Message = $"Procesando hoja '{sheetName}'..."
                    });

                    var hasBase = baseSheets.TryGetValue(sheetName, out var baseSheet);
                    var hasNew = newSheets.TryGetValue(sheetName, out var newSheet);

                    if (!hasBase && hasNew)
                    {
                        // Hoja nueva completa: copiar tal cual (no hay cambios respecto a base)
                        newSheet.CopyTo(outWb, sheetName);
                        continue;
                    }
                    if (hasBase && !hasNew)
                    {
                        // Hoja desaparecida: copiar tal cual del base (sin cambios)
                        baseSheet.CopyTo(outWb, sheetName);
                        continue;
                    }

                    // Comparar filas entre base y nuevo. Se asume el mismo formato que exportación:
                    // Encabezados en la fila 1, columnas:
                    // 1:Fichero 2:URL 3:Usuario 4:Carpeta 5:Software 6:Emails 7:Clientes (equipos)
                    // Clave de identidad: URL (si no hay, usar Fichero+URL)
                    var result = BuildComparisonSheet(outWb, sheetName, baseSheet, newSheet, keyMode, highlight);
                }

                report?.Invoke(new CompareProgress { CurrentStep = total + 1, TotalSteps = total + 1, Message = "Guardando informe..." });
                // Si el destino existe, eliminar para sobrescribir sin prompt
                try { if (System.IO.File.Exists(outputPath)) System.IO.File.Delete(outputPath); } catch { }
                outWb.SaveAs(outputPath);
            }

            report?.Invoke(new CompareProgress { CurrentStep = 100, TotalSteps = 100, Message = "Completado" });
        }

        private static IXLWorksheet BuildComparisonSheet(XLWorkbook outWb, string sheetName, IXLWorksheet baseSheet, IXLWorksheet newSheet, string keyMode, bool highlight)
        {
            var ws = outWb.Worksheets.Add(sheetName);

            // Copiar cabeceras con columnas extra para estado/cambios
            ws.Cell(1, 1).Value = "Fichero";
            ws.Cell(1, 2).Value = "URL";
            ws.Cell(1, 3).Value = "Usuario";
            ws.Cell(1, 4).Value = "Carpeta";
            ws.Cell(1, 5).Value = "Software";
            ws.Cell(1, 6).Value = "Emails";
            ws.Cell(1, 7).Value = "Clientes (equipos)";
            ws.Cell(1, 8).Value = "Estado"; // Nuevo, Eliminado, Cambiado, Igual
            ws.Cell(1, 9).Value = "Detalle cambios"; // campos cambiados

            var baseRows = ReadRows(baseSheet, keyMode);
            var newRows = ReadRows(newSheet, keyMode);

            var baseMap = baseRows.ToDictionary(r => r.Key, r => r, StringComparer.OrdinalIgnoreCase);
            var newMap = newRows.ToDictionary(r => r.Key, r => r, StringComparer.OrdinalIgnoreCase);

            var keys = new HashSet<string>(baseMap.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var k in newMap.Keys) keys.Add(k);

            int row = 2;
            var states = new List<(int Row, string Estado)>();
            foreach (var key in keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase))
            {
                var hasB = baseMap.TryGetValue(key, out var b);
                var hasN = newMap.TryGetValue(key, out var n);

                if (hasB && !hasN)
                {
                    // Eliminado
                    WriteRow(ws, row, b, "Eliminado", "");
                    if (highlight) states.Add((row, "Eliminado"));
                    row++;
                }
                else if (!hasB && hasN)
                {
                    // Nuevo
                    WriteRow(ws, row, n, "Nuevo", "");
                    if (highlight) states.Add((row, "Nuevo"));
                    row++;
                }
                else
                {
                    // Comparar campos
                    var diffs = DiffFields(b, n);
                    var estado = diffs.Count == 0 ? "Igual" : "Cambiado";
                    var detalle = string.Join(", ", diffs);
                    // Escribir los valores «nuevos» para reflejar el estado actual
                    WriteRow(ws, row, n, estado, detalle);
                    if (highlight) states.Add((row, estado));
                    row++;
                }
            }

            // Tabla y formato
            var totalRows = row - 1;
            var range = ws.Range(1, 1, Math.Max(1, totalRows), 9);
            var table = range.CreateTable();
            table.ShowAutoFilter = true;
            table.Theme = XLTableTheme.TableStyleMedium9;
            ws.SheetView.FreezeRows(1);
            // Anchos similares al exportador
            double PxToWidth(int px) => Math.Max(1, (px - 5) / 7.0);
            ws.Column(1).Width = PxToWidth(400);
            ws.Column(2).Width = PxToWidth(600);
            ws.Column(3).Width = PxToWidth(400);
            ws.Column(4).Width = PxToWidth(400);
            ws.Column(5).Width = PxToWidth(400);
            ws.Column(6).Width = PxToWidth(300);
            ws.Column(7).Width = PxToWidth(300);
            ws.Column(8).Width = PxToWidth(150);
            ws.Column(9).Width = PxToWidth(300);

            // Ajuste de texto y alineación como en exportación
            if (totalRows >= 1)
            {
                var used = ws.Range(1, 1, totalRows, 9);
                used.Style.Alignment.WrapText = true;
                used.Style.Alignment.ShrinkToFit = false;
                used.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            // Reaplicar resaltado después de crear la tabla para que no lo sobrescriba el estilo
            foreach (var s in states)
            {
                Highlight(ws, s.Row, s.Estado);
            }
            return ws;
        }

        private static void Highlight(IXLWorksheet ws, int row, string estado)
        {
            var range = ws.Range(row, 1, row, 9);
            switch (estado)
            {
                case "Nuevo":
                    range.Style.Fill.BackgroundColor = XLColor.LightGreen; // verde flojo
                    break;
                case "Eliminado":
                    range.Style.Fill.BackgroundColor = XLColor.LightSalmon; // rojo flojo
                    break;
                case "Cambiado":
                    // No colorear explícitamente; se mantiene el estilo de tabla
                    break;
                case "Igual":
                    range.Style.Fill.BackgroundColor = XLColor.Gainsboro; // gris flojo
                    break;
                default:
                    break;
            }
        }

        private static List<string> DiffFields(RowData b, RowData n)
        {
            var diffs = new List<string>();
            if (!StringEquals(b.FileName, n.FileName)) diffs.Add("Fichero");
            if (!StringEquals(b.Url, n.Url)) diffs.Add("URL");
            if (!StringEquals(NormalizeMultiline(b.User), NormalizeMultiline(n.User))) diffs.Add("Usuario");
            if (!StringEquals(NormalizeMultiline(b.Folder), NormalizeMultiline(n.Folder))) diffs.Add("Carpeta");
            if (!StringEquals(NormalizeMultiline(b.Software), NormalizeMultiline(n.Software))) diffs.Add("Software");
            if (!StringEquals(NormalizeMultiline(b.Emails), NormalizeMultiline(n.Emails))) diffs.Add("Emails");
            if (!StringEquals(NormalizeMultiline(b.Clients), NormalizeMultiline(n.Clients))) diffs.Add("Clientes (equipos)");
            return diffs;
        }

        private static string NormalizeMultiline(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            // Normalizar separadores y ordenar líneas para reducir ruido
            var lines = value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase);
            return string.Join("\n", lines);
        }

        private static bool StringEquals(string a, string b)
        {
            return string.Equals(a ?? string.Empty, b ?? string.Empty, StringComparison.Ordinal);
        }

        private static void WriteRow(IXLWorksheet ws, int row, RowData d, string estado, string detalle)
        {
            ws.Cell(row, 1).Value = d.FileName;
            ws.Cell(row, 2).Value = d.Url;
            ws.Cell(row, 3).Value = d.User;
            ws.Cell(row, 4).Value = d.Folder;
            ws.Cell(row, 5).Value = d.Software;
            ws.Cell(row, 6).Value = d.Emails;
            ws.Cell(row, 7).Value = d.Clients;
            ws.Cell(row, 8).Value = estado;
            ws.Cell(row, 9).Value = detalle;
        }

        private static IEnumerable<RowData> ReadRows(IXLWorksheet sheet, string keyMode)
        {
            var used = sheet.RangeUsed();
            if (used == null)
                yield break;

            var maxRow = used.RowCount();
            for (int r = 2; r <= maxRow; r++)
            {
                var d = new RowData
                {
                    FileName = sheet.Cell(r, 1).GetString(),
                    Url = sheet.Cell(r, 2).GetString(),
                    User = sheet.Cell(r, 3).GetString(),
                    Folder = sheet.Cell(r, 4).GetString(),
                    Software = sheet.Cell(r, 5).GetString(),
                    Emails = sheet.Cell(r, 6).GetString(),
                    Clients = sheet.Cell(r, 7).GetString()
                };
                if (string.Equals(keyMode, "Fichero|URL", StringComparison.OrdinalIgnoreCase))
                {
                    d.Key = ($"{d.FileName}|{d.Url}");
                }
                else
                {
                    // Clave: URL si existe, si no Fichero|URL
                    d.Key = string.IsNullOrWhiteSpace(d.Url) ? ($"{d.FileName}|{d.Url}") : d.Url;
                }
                yield return d;
            }
        }

        private sealed class RowData
        {
            public string Key { get; set; }
            public string FileName { get; set; }
            public string Url { get; set; }
            public string User { get; set; }
            public string Folder { get; set; }
            public string Software { get; set; }
            public string Emails { get; set; }
            public string Clients { get; set; }
        }
    }
}


