using ClosedXML.Excel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using FocaExcelExport.Classes;
using System.Net;
using System.Text.RegularExpressions;
using System.Text;

namespace FocaExcelExport
{
    // Class to track export progress
        public class ExportProgress
    {
        public int CurrentRecord { get; set; }
        public int TotalRecords { get; set; }
        public string StatusMessage { get; set; }
        public int PercentComplete 
        { 
            get 
            { 
                if (TotalRecords <= 0) return 0;
                var pct = (int)Math.Round((double)CurrentRecord * 100.0 / TotalRecords);
                if (pct < 0) pct = 0;
                if (pct > 100) pct = 100;
                return pct; 
            } 
        }
    }

    public class Exporter
    {
        private readonly string _connectionString;
        
        public Exporter(string connectionString)
        {
            _connectionString = connectionString;
        }

        public async Task ExportToExcelAsync(int projectId, string fileName, IProgress<ExportProgress> progress = null)
        {
            try
            {
                // Discover schema
                var schemaResolver = new SchemaResolver(_connectionString);
                
                var projectsTable = await schemaResolver.FindProjectsTableAsync();
                var filesTable = await schemaResolver.FindFilesTableAsync();
                var metadataTable = await schemaResolver.FindMetadataTableAsync();
                var emailItemsTable = await schemaResolver.FindEmailItemsTableAsync();
                var userItemsTable = await schemaResolver.FindUserItemsTableAsync();
                var applicationsTable = await schemaResolver.FindApplicationsTableAsync();
                var applicationItemsTable = await schemaResolver.FindApplicationItemsTableAsync();
                var serversTable = await schemaResolver.FindServersTableAsync();
                var serverItemsTable = await schemaResolver.FindServerItemsTableAsync();
                var computersItemsTable = await schemaResolver.FindComputersItemsTableAsync();

                // Find relevant columns
                var projectPkColumn = await schemaResolver.FindProjectIdColumnAsync(projectsTable);
                var filePkColumn = await schemaResolver.FindFileIdColumnAsync(filesTable);
                var filesProjectFkColumn = await schemaResolver.FindFilesProjectFkColumnAsync(filesTable);
                
                var fileNameColumn = await schemaResolver.FindFileNameColumnAsync(filesTable);
                var urlColumn = await schemaResolver.FindUrlColumnAsync(filesTable);
                
                var userNameColumn = await schemaResolver.FindUserNameColumnAsync(metadataTable);
                var locationColumn = await schemaResolver.FindLocationColumnAsync(metadataTable);
                var emailColumn = await schemaResolver.FindEmailColumnAsync(metadataTable);
                var clientColumn = await schemaResolver.FindClientColumnAsync(projectsTable);

                // MetaExtractors available columns
                var metaCols = await schemaResolver.GetColumnsAsync("MetaExtractors");
                bool hasFoundUsers = metaCols.Contains("FoundUsers_Id");
                bool hasFoundEmails = metaCols.Contains("FoundEmails_Id");
                bool hasFoundApplications = false; // aplicaciones vienen vía md.Applications_Id
                bool hasFoundServers = metaCols.Contains("FoundServers_Id");

                // Count total records for progress
                var totalRecords = await GetRecordCountAsync(projectId, projectsTable, filesTable, metadataTable, 
                    projectPkColumn, filePkColumn, filesProjectFkColumn);
                
                int currentRecord = 0;
                progress?.Report(new ExportProgress
                {
                    CurrentRecord = currentRecord,
                    TotalRecords = totalRecords,
                    StatusMessage = "Iniciando exportación..."
                });

                // Build dynamic SQL query
                var query = BuildExportQuery(projectsTable, filesTable, metadataTable, 
                    projectPkColumn, filePkColumn, filesProjectFkColumn, fileNameColumn, urlColumn, 
                    userNameColumn, locationColumn, emailColumn, clientColumn, emailItemsTable, userItemsTable,
                    applicationsTable, applicationItemsTable, serversTable, serverItemsTable,
                    computersItemsTable,
                    hasFoundUsers, hasFoundEmails, hasFoundApplications, hasFoundServers);

                // Resolve worksheet name from project
                var worksheetName = await GetProjectNameAsync(projectId, projectsTable, projectPkColumn);
                if (string.IsNullOrWhiteSpace(worksheetName)) worksheetName = "Export";
                worksheetName = SanitizeWorksheetName(worksheetName);

                // Create Excel file
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(worksheetName);
                    
                    // Add headers (según captura: Fichero, URL, Usuario, Carpeta, Software, Emails, Clientes (equipos))
                    worksheet.Cell(1, 1).Value = "Fichero";
                    worksheet.Cell(1, 2).Value = "URL";
                    worksheet.Cell(1, 3).Value = "Usuario";
                    worksheet.Cell(1, 4).Value = "Carpeta";
                    worksheet.Cell(1, 5).Value = "Software";
                    worksheet.Cell(1, 6).Value = "Emails";
                    worksheet.Cell(1, 7).Value = "Clientes (equipos)";
                    worksheet.Cell(1, 8).Value = "Fecha de creacion";
                    worksheet.Cell(1, 9).Value = "Fecha de modificacion";

                    int row = 2; // Start from row 2 as row 1 has headers
                    
                    using (var connection = new SqlConnection(_connectionString))
                    {
                        await connection.OpenAsync();
                        
                        using (var command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@ProjectId", projectId);
                            using (var reader = await command.ExecuteReaderAsync())
                            {
                                while (await reader.ReadAsync())
                                {
                                    // Write data to Excel
                                    var urlValue = reader[1]?.ToString() ?? "";
                                    var computedName = reader[0]?.ToString() ?? "";
                                    var info = await TryResolveHttpInfoFromUrlAsync(urlValue);
                                    var resolvedName = info?.FileName ?? computedName;
                                    worksheet.Cell(row, 1).Value = resolvedName;
                                    worksheet.Cell(row, 2).Value = urlValue;
                                    worksheet.Cell(row, 3).Value = reader[2]?.ToString() ?? "";
                                    worksheet.Cell(row, 4).Value = reader[3]?.ToString() ?? "";
                                    worksheet.Cell(row, 5).Value = reader[4]?.ToString() ?? "";
                                    worksheet.Cell(row, 6).Value = reader[5]?.ToString() ?? "";
                                    worksheet.Cell(row, 7).Value = reader[6]?.ToString() ?? "";
                                    // Fechas PDF si están disponibles
                                    var createdStr = info?.PdfCreationUtc.HasValue == true ? info.PdfCreationUtc.Value.ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss") : string.Empty;
                                    var modifiedStr = info?.PdfModifiedUtc.HasValue == true ? info.PdfModifiedUtc.Value.ToLocalTime().ToString("dd/MM/yyyy HH:mm:ss") : string.Empty;
                                    worksheet.Cell(row, 8).Value = createdStr;
                                    worksheet.Cell(row, 9).Value = modifiedStr;

                                    currentRecord++;
                                    row++;

                                    // Report progress
                                    if (progress != null)
                                    {
                                        var safeTotal = totalRecords;
                                        if (safeTotal < currentRecord) safeTotal = currentRecord;
                                        progress.Report(new ExportProgress
                                        {
                                            CurrentRecord = currentRecord,
                                            TotalRecords = safeTotal,
                                            StatusMessage = $"Procesando registro {currentRecord} de {safeTotal}..."
                                        });
                                    }
                                }
                            }
                        }
                    }

					// Format header row (sin forzar color de fondo)
					var headerRange = worksheet.Range(1, 1, 1, 9);
					headerRange.Style.Font.Bold = true;

                    // Convert data range to an Excel Table with headers
                    var totalRows = row - 1;
                    if (totalRows >= 1)
                    {
                        var tableRange = worksheet.Range(1, 1, totalRows, 9);
                        var table = tableRange.CreateTable();
                        table.ShowAutoFilter = true;
                        table.Theme = XLTableTheme.TableStyleMedium9;
                    }
                    worksheet.SheetView.FreezeRows(1);

                    // Set explicit column widths (approximate pixels)
                    // ClosedXML uses Excel units (character width). Approximate conversion: px ≈ (width*7)+5
                    // So width ≈ (px-5)/7
                    double PxToWidth(int px) => Math.Max(1, (px - 5) / 7.0);
                    worksheet.Column(1).Width = PxToWidth(400); // Fichero
                    worksheet.Column(2).Width = PxToWidth(600); // URL
                    worksheet.Column(3).Width = PxToWidth(400); // Usuario
                    worksheet.Column(4).Width = PxToWidth(400); // Carpeta
                    worksheet.Column(5).Width = PxToWidth(400); // Software
                    worksheet.Column(6).Width = PxToWidth(300); // Emails
                    worksheet.Column(7).Width = PxToWidth(300); // Clientes (equipos)
                    worksheet.Column(8).Width = PxToWidth(220); // Fecha de creacion
                    worksheet.Column(9).Width = PxToWidth(220); // Fecha de modificacion
                    // Evitar que el texto desborde a la siguiente celda: no envolver y ajustar para encajar
                    var usedRange = worksheet.Range(1, 1, totalRows, 9);
                    // Ajustar texto en cada celda (sin reducir tipografía)
                    usedRange.Style.Alignment.WrapText = true;
                    usedRange.Style.Alignment.ShrinkToFit = false;
                    usedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    usedRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // Save the workbook
                    workbook.SaveAs(fileName);
                }

                progress?.Report(new ExportProgress 
                { 
                    CurrentRecord = totalRecords, 
                    TotalRecords = totalRecords, 
                    StatusMessage = "Export completed successfully!" 
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Export failed: {ex.Message}", ex);
            }
        }

        private async Task<int> GetRecordCountAsync(int projectId, string projectsTable, string filesTable, string metadataTable,
            string projectPkColumn, string filePkColumn, string filesProjectFkColumn)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                
                // Build count query
                string countQuery;
                
                if (!string.IsNullOrEmpty(metadataTable))
                {
                    // Include metadata table in the query
                    countQuery = $@"
                        SELECT COUNT(*)
                        FROM [dbo].[{filesTable}] f
                        JOIN [dbo].[{projectsTable}] p ON f.[{filesProjectFkColumn}] = p.[{projectPkColumn}]
                        LEFT JOIN [dbo].[{metadataTable}] m ON m.[{filePkColumn}] = f.[{filePkColumn}]
                        WHERE p.[{projectPkColumn}] = @ProjectId";
                }
                else
                {
                    // Exclude metadata table
                    countQuery = $@"
                        SELECT COUNT(*)
                        FROM [dbo].[{filesTable}] f
                        JOIN [dbo].[{projectsTable}] p ON f.[{filesProjectFkColumn}] = p.[{projectPkColumn}]
                        WHERE p.[{projectPkColumn}] = @ProjectId";
                }
                
                using (var command = new SqlCommand(countQuery, connection))
                {
                    command.Parameters.AddWithValue("@ProjectId", projectId);
                    var result = await command.ExecuteScalarAsync();
                    return Convert.ToInt32(result);
                }
            }
        }

        private async Task<string> GetProjectNameAsync(int projectId, string projectsTable, string projectPkColumn)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var sql = $"SELECT TOP 1 ProjectName FROM [dbo].[{projectsTable}] WHERE [{projectPkColumn}] = @id";
                using (var cmd = new SqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@id", projectId);
                    var obj = await cmd.ExecuteScalarAsync();
                    return obj?.ToString();
                }
            }
        }

        private static string SanitizeWorksheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Export";
            var invalid = new char[] { '\\', '/', '*', '[', ']', ':', '?' };
            foreach (var c in invalid) name = name.Replace(c, '_');
            if (name.Length > 31) name = name.Substring(0, 31);
            return name.Trim();
        }

        private string BuildExportQuery(string projectsTable, string filesTable, string metadataTable,
            string projectPkColumn, string filePkColumn, string filesProjectFkColumn, string fileNameColumn, string urlColumn,
            string userNameColumn, string locationColumn, string emailColumn, string clientColumn, string emailItemsTable, string userItemsTable,
            string applicationsTable, string applicationItemsTable, string serversTable, string serverItemsTable,
            string computersItemsTable,
            bool hasFoundUsers, bool hasFoundEmails, bool hasFoundApplications, bool hasFoundServers)
        {
            // Based on FOCA migration structure and corrections:
            // Projects table: Id, ProjectName, Domain
            // FilesITems table: Id, IdProject(FK), URL, Path, Metadata_Id(FK) 
            // FilesITems.Metadata_Id → MetaDatas.Id (NOT MetaExtractors directly)
            // MetaExtractors table: Id, with FKs like FoundUsers_Id, FoundEmails_Id, FoundMetaData_Id
            // MetaDatas table: Id, with FKs to MetaExtractors via FoundMetaData_Id
            // UserItems table: Id, Name, Users_Id (FK to Users)
            // EmailItems table: Id, Mail, Emails_Id (FK to Emails)
            
            // For Fichero (file name), compute base name from Path or URL
            // Windows path base name: RIGHT(Path, CHARINDEX('\\\
            //', REVERSE(Path)+'\\')-1)
            // URL base name: RIGHT(URL, CHARINDEX('/', REVERSE(URL) + '/') - 1)
            // "Fichero" será recalculado a partir de la URL (Content-Disposition) en tiempo de exportación
            string ficheroCol = "RIGHT(f.[URL], CHARINDEX('/', REVERSE(f.[URL]) + '/') - 1)";
            // For URL, use the URL column
            string urlCol = "f.[URL]";
            // 'Carpeta' se agregará desde [PathsItems] en la selección final

            string query;
            var userItemsJoinTable = string.IsNullOrEmpty(userItemsTable) ? "UserItems" : userItemsTable;
            var emailItemsJoinTable = string.IsNullOrEmpty(emailItemsTable) ? "EmailItems" : emailItemsTable;
            var applicationsJoinTable = string.IsNullOrEmpty(applicationsTable) ? "Applications" : applicationsTable;
            var applicationItemsJoinTable = string.IsNullOrEmpty(applicationItemsTable) ? "ApplicationItems" : applicationItemsTable;
            var serversJoinTable = string.IsNullOrEmpty(serversTable) ? "Servers" : serversTable;
            var serverItemsJoinTable = string.IsNullOrEmpty(serverItemsTable) ? "ServerItems" : serverItemsTable;

            bool enableUsers = hasFoundUsers && !string.IsNullOrEmpty(userItemsJoinTable);
            bool enableEmails = hasFoundEmails && !string.IsNullOrEmpty(emailItemsJoinTable);
            bool enableApplications = !string.IsNullOrEmpty(applicationsJoinTable) && !string.IsNullOrEmpty(applicationItemsJoinTable);
            bool enableServers = hasFoundServers && !string.IsNullOrEmpty(serversJoinTable) && !string.IsNullOrEmpty(serverItemsJoinTable);
            bool enableComputers = !string.IsNullOrEmpty(computersItemsTable);
            bool enableComputersByUserId = enableComputers && hasFoundUsers; // requiere Users_Id desde metadatos
            bool enableComputersByPcUsername = enableComputers && hasFoundUsers; // usar UserItems con IsComputerUser=1

            // Column expressions depend on whether we enable the OUTER APPLY blocks
            string usuarioCol = enableUsers ? "ISNULL(ou.Usuario, '')" : "''";
            string emailCol = enableEmails ? "ISNULL(oe.Email, '')" : "''";
            string softwareCol = enableApplications ? "ISNULL(osw.Software, '')" : "''";
            string equiposCol = enableComputersByPcUsername ? "ISNULL(ocn.Equipo, ISNULL(ocx.Equipo, ''))" : (enableComputersByUserId ? "ISNULL(ocx.Equipo, '')" : (enableComputers ? "''" : "''"));

            // Agregar metadatos en columnas (una fila por fichero)
            string meUsersIdExpr = "ISNULL(me_direct.[FoundUsers_Id], me.[FoundUsers_Id])";
            string meEmailsIdExpr = "ISNULL(me_direct.[FoundEmails_Id], me.[FoundEmails_Id])";

            string aggUsersExpr = enableUsers ? $@"ISNULL(STUFF((SELECT DISTINCT CHAR(10) + ui.Name
                FROM [dbo].[Users] u JOIN [dbo].[{userItemsJoinTable}] ui ON ui.[Users_Id] = u.[Id]
                WHERE u.[Id] = {meUsersIdExpr}
                FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,1,''), '')" : "''";

            string aggEmailsExpr = enableEmails ? $@"ISNULL(STUFF((SELECT DISTINCT CHAR(10) + ei.Mail
                FROM [dbo].[Emails] e JOIN [dbo].[{emailItemsJoinTable}] ei ON ei.[Emails_Id] = e.[Id]
                WHERE e.[Id] = {meEmailsIdExpr}
                FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,1,''), '')" : "''";

            string aggSoftwareExpr = enableApplications ? $@"ISNULL(STUFF((SELECT DISTINCT CHAR(10) + ai.Name
                FROM [dbo].[{applicationItemsJoinTable}] ai
                WHERE ai.[Applications_Id] = md.[Applications_Id]
                FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,1,''), '')" : "''";

            string aggFoldersExpr = $@"ISNULL(STUFF((SELECT DISTINCT CHAR(10) + pi.[Path]
                FROM [dbo].[Paths] pth JOIN [dbo].[PathsItems] pi ON pi.[Paths_Id] = pth.[Id]
                WHERE pth.[Id] = ISNULL(me_direct.[FoundPaths_Id], me.[FoundPaths_Id]) AND pi.[Path] IS NOT NULL
                FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,1,''), '')";

            string aggClientsExpr = enableComputers ? $@"ISNULL(STUFF((SELECT DISTINCT CHAR(10) + ci.name
                FROM [dbo].[{computersItemsTable}] ci
                WHERE ci.[IdProject] = f.[{filesProjectFkColumn}] AND ci.[type] = 0 AND (
                    ci.[Users_Id] = {meUsersIdExpr} OR ci.[RemoteUsers_Id] = {meUsersIdExpr}
                    OR UPPER(REPLACE(ci.[name],' ','')) IN (
                        SELECT UPPER('PC_'+REPLACE(ui2.[Name],' ',''))
                        FROM [dbo].[Users] u2 JOIN [dbo].[{userItemsJoinTable}] ui2 ON ui2.[Users_Id] = u2.[Id]
                        WHERE u2.[Id] = {meUsersIdExpr}
                    )
                )
                FOR XML PATH(''), TYPE).value('.','nvarchar(max)'),1,1,''), '')" : "''";

            query = $@"
                SELECT
                    {ficheroCol} AS [Fichero],
                    {urlCol} AS [URL],
                    {aggUsersExpr} AS [Usuario],
                    {aggFoldersExpr} AS [Carpeta],
                    {aggSoftwareExpr} AS [Software],
                    {aggEmailsExpr} AS [Emails],
                    {aggClientsExpr} AS [Clientes (equipos)]
                FROM [dbo].[{filesTable}] f
                JOIN [dbo].[{projectsTable}] p ON f.[{filesProjectFkColumn}] = p.[{projectPkColumn}]
                LEFT JOIN [dbo].[MetaExtractors] me_direct ON f.[Metadata_Id] = me_direct.[Id]
                LEFT JOIN [dbo].[MetaDatas] md ON f.[Metadata_Id] = md.[Id]
                LEFT JOIN [dbo].[MetaExtractors] me ON me.[FoundMetaData_Id] = md.[Id]
                WHERE p.[{projectPkColumn}] = @ProjectId
                ORDER BY {ficheroCol}";

            return query;
        }

        private static string GetFileNameFromContentDisposition(string contentDisposition)
        {
            if (string.IsNullOrWhiteSpace(contentDisposition)) return null;
            try
            {
                // RFC 5987 filename*
                var match = Regex.Match(contentDisposition, @"filename\*=UTF-8''([^;]+)", RegexOptions.IgnoreCase);
                if (match.Success) return Uri.UnescapeDataString(match.Groups[1].Value);
                // filename="name.ext" o filename=name.ext
                match = Regex.Match(contentDisposition, @"filename=""?([^"";]+)""?", RegexOptions.IgnoreCase);
                if (match.Success) return match.Groups[1].Value;
            }
            catch { }
            return null;
        }

        // Información resuelta por HTTP y metadatos PDF
        private sealed class ResolvedHttpInfo
        {
            public string FileName { get; set; }
            public DateTime? PdfCreationUtc { get; set; }
            public DateTime? PdfModifiedUtc { get; set; }
        }

        private static async Task<ResolvedHttpInfo> TryResolveHttpInfoFromUrlAsync(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return null;

            string fileName = null;
            string contentType = null;
            long contentLength = -1;

            // HEAD primero
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "HEAD";
                req.Timeout = 5000;
                req.UserAgent = "FocaExcelExport/1.0";
                using (var resp = (HttpWebResponse)await req.GetResponseAsync())
                {
                    fileName = GetFileNameFromContentDisposition(resp.Headers["Content-Disposition"]) ?? fileName;
                    contentType = resp.ContentType;
                    long len;
                    if (long.TryParse(resp.Headers["Content-Length"], out len)) contentLength = len;
                }
            }
            catch { }

            // Si aún no hay nombre, intentar con GET parcial (también útil para servidores que no soportan HEAD)
            if (string.IsNullOrWhiteSpace(fileName))
            {
                try
                {
                    var req = (HttpWebRequest)WebRequest.Create(url);
                    req.Method = "GET";
                    req.Timeout = 5000;
                    req.UserAgent = "FocaExcelExport/1.0";
                    req.AddRange(0, 0);
                    using (var resp = (HttpWebResponse)await req.GetResponseAsync())
                    {
                        fileName = GetFileNameFromContentDisposition(resp.Headers["Content-Disposition"]) ?? fileName;
                        if (string.IsNullOrEmpty(contentType)) contentType = resp.ContentType;
                        if (contentLength < 0) contentLength = resp.ContentLength;
                    }
                }
                catch { }
            }

            // Último segmento de URL como fallback del nombre
            if (string.IsNullOrWhiteSpace(fileName))
            {
                try
                {
                    var uri = new Uri(url, UriKind.RelativeOrAbsolute);
                    var last = uri.IsAbsoluteUri ? uri.Segments[uri.Segments.Length - 1] : url.Substring(url.LastIndexOf('/') + 1);
                    if (!string.IsNullOrWhiteSpace(last)) fileName = last;
                }
                catch { }
            }

            var info = new ResolvedHttpInfo { FileName = fileName };

            // Extraer fechas del PDF si procede
            bool looksPdf = false;
            try
            {
                if (!string.IsNullOrEmpty(contentType) && contentType.IndexOf("pdf", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    looksPdf = true;
                }
                else
                {
                    // Heurística por extensión de la URL
                    var cleanUrl = url;
                    int q = cleanUrl.IndexOf('?');
                    if (q >= 0) cleanUrl = cleanUrl.Substring(0, q);
                    int h = cleanUrl.IndexOf('#');
                    if (h >= 0) cleanUrl = cleanUrl.Substring(0, h);
                    looksPdf = cleanUrl.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch { }

            if (looksPdf)
            {
                try
                {
                    var dates = await TryFetchPdfDatesAsync(url, contentLength);
                    if (dates != null)
                    {
                        info.PdfCreationUtc = dates.Item1;
                        info.PdfModifiedUtc = dates.Item2;
                    }
                }
                catch { }
            }

            return info;
        }

        private static async Task<Tuple<DateTime?, DateTime?>> TryFetchPdfDatesAsync(string url, long contentLength)
        {
            // Leer primeros bytes (XMP suele estar al principio)
            string headText = null;
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Timeout = 7000;
                req.UserAgent = "FocaExcelExport/1.0";
                req.AddRange(0, (int)Math.Min(1024 * 1024 - 1, Math.Max(0, contentLength > 0 ? Math.Min(contentLength - 1, 1024 * 1024 - 1) : 1024 * 1024 - 1)));
                using (var resp = (HttpWebResponse)await req.GetResponseAsync())
                using (var ms = new MemoryStream())
                {
                    await resp.GetResponseStream().CopyToAsync(ms);
                    var bytes = ms.ToArray();
                    headText = Encoding.ASCII.GetString(bytes);
                }
            }
            catch { }

            DateTime? created = null;
            DateTime? modified = null;

            if (!string.IsNullOrEmpty(headText))
            {
                // XMP ISO8601
                try
                {
                    var m1 = Regex.Match(headText, @"<xmp:CreateDate>([^<]+)</xmp:CreateDate>", RegexOptions.IgnoreCase);
                    if (m1.Success) created = TryParseIsoOrPdfDate(m1.Groups[1].Value);
                    var m2 = Regex.Match(headText, @"<xmp:ModifyDate>([^<]+)</xmp:ModifyDate>", RegexOptions.IgnoreCase);
                    if (m2.Success) modified = TryParseIsoOrPdfDate(m2.Groups[1].Value);
                }
                catch { }
            }

            // Si no las encontramos, intentar en la cola (Info dictionary suele ir al final)
            if ((!created.HasValue || !modified.HasValue) && contentLength > 0)
            {
                try
                {
                    long tail = Math.Min(256 * 1024, contentLength);
                    long from = Math.Max(0, contentLength - tail);
                    var req = (HttpWebRequest)WebRequest.Create(url);
                    req.Method = "GET";
                    req.Timeout = 7000;
                    req.UserAgent = "FocaExcelExport/1.0";
                    req.AddRange(from, contentLength - 1);
                    using (var resp = (HttpWebResponse)await req.GetResponseAsync())
                    using (var ms = new MemoryStream())
                    {
                        await resp.GetResponseStream().CopyToAsync(ms);
                        var bytes = ms.ToArray();
                        var tailText = Encoding.ASCII.GetString(bytes);
                        if (!created.HasValue)
                        {
                            var m = Regex.Match(tailText, @"/CreationDate\s*\(([^\)]+)\)");
                            if (m.Success) created = TryParseIsoOrPdfDate(m.Groups[1].Value);
                        }
                        if (!modified.HasValue)
                        {
                            var m = Regex.Match(tailText, @"/ModDate\s*\(([^\)]+)\)");
                            if (m.Success) modified = TryParseIsoOrPdfDate(m.Groups[1].Value);
                        }
                    }
                }
                catch { }
            }

            return Tuple.Create(created, modified);
        }

        private static DateTime? TryParseIsoOrPdfDate(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            try
            {
                // Quitar prefijo D:
                var s = raw.Trim();
                if (s.StartsWith("D:", StringComparison.OrdinalIgnoreCase)) s = s.Substring(2);

                // Intentar ISO 8601 primero
                if (DateTimeOffset.TryParse(s, out var dto))
                {
                    return dto.UtcDateTime;
                }

                // Formato PDF: YYYYMMDDHHmmSSOHH'mm
                var re = new Regex(@"^(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?([Zz]|[\+\-]\d{2}'?\d{2}')?", RegexOptions.Compiled);
                var m = re.Match(s);
                if (!m.Success) return null;

                int year = int.Parse(m.Groups[1].Value);
                int month = string.IsNullOrEmpty(m.Groups[2].Value) ? 1 : int.Parse(m.Groups[2].Value);
                int day = string.IsNullOrEmpty(m.Groups[3].Value) ? 1 : int.Parse(m.Groups[3].Value);
                int hour = string.IsNullOrEmpty(m.Groups[4].Value) ? 0 : int.Parse(m.Groups[4].Value);
                int minute = string.IsNullOrEmpty(m.Groups[5].Value) ? 0 : int.Parse(m.Groups[5].Value);
                int second = string.IsNullOrEmpty(m.Groups[6].Value) ? 0 : int.Parse(m.Groups[6].Value);
                var offsetStr = m.Groups[7].Value;

                TimeSpan offset = TimeSpan.Zero;
                if (!string.IsNullOrEmpty(offsetStr) && !offsetStr.Equals("Z", StringComparison.OrdinalIgnoreCase))
                {
                    // +HH'mm o -HH'mm o +HHmm
                    var sign = offsetStr[0] == '-' ? -1 : 1;
                    var digits = offsetStr.TrimStart('+', '-').Replace("'", string.Empty);
                    int oh = 0, om = 0;
                    if (digits.Length >= 2) oh = int.Parse(digits.Substring(0, 2));
                    if (digits.Length >= 4) om = int.Parse(digits.Substring(2, 2));
                    offset = new TimeSpan(sign * oh, sign * om, 0);
                }

                var dto2 = new DateTimeOffset(year, month, day, hour, minute, second, offset);
                return dto2.UtcDateTime;
            }
            catch { return null; }
        }
    }
}