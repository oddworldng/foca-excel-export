using ClosedXML.Excel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using FocaExcelExport.Classes;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;

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
        private static readonly ConcurrentDictionary<string, string> _urlToFileNameCache = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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

                // Create Excel file
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Exported Data");
                    
                    // Add headers (según captura: Fichero, URL, Usuario, Carpeta, Software, Emails, Clientes (equipos))
                    worksheet.Cell(1, 1).Value = "Fichero";
                    worksheet.Cell(1, 2).Value = "URL";
                    worksheet.Cell(1, 3).Value = "Usuario";
                    worksheet.Cell(1, 4).Value = "Carpeta";
                    worksheet.Cell(1, 5).Value = "Software";
                    worksheet.Cell(1, 6).Value = "Emails";
                    worksheet.Cell(1, 7).Value = "Clientes (equipos)";

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
                                    var resolvedName = await TryResolveFileNameFromUrlAsync(urlValue) ?? computedName;
                                    worksheet.Cell(row, 1).Value = resolvedName;
                                    worksheet.Cell(row, 2).Value = urlValue;
                                    worksheet.Cell(row, 3).Value = reader[2]?.ToString() ?? "";
                                    worksheet.Cell(row, 4).Value = reader[3]?.ToString() ?? "";
                                    worksheet.Cell(row, 5).Value = reader[4]?.ToString() ?? "";
                                    worksheet.Cell(row, 6).Value = reader[5]?.ToString() ?? "";
                                    worksheet.Cell(row, 7).Value = reader[6]?.ToString() ?? "";

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

                    // Format header row
                    var headerRange = worksheet.Range(1, 1, 1, 7);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                    // Convert data range to an Excel Table with headers
                    var totalRows = row - 1;
                    if (totalRows >= 1)
                    {
                        var tableRange = worksheet.Range(1, 1, totalRows, 7);
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
                    // Evitar que el texto desborde a la siguiente celda: no envolver y ajustar para encajar
                    var usedRange = worksheet.Range(1, 1, totalRows, 7);
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
            // Para 'Carpeta' queremos Folders desde [PathsItems]; en las filas no-Folders irá vacío
            string ubicacionCol = "''";

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
            
			// Construir una única hoja con posibles filas repetidas por fichero (una por cada valor de metadato)
			// Base FROM (sin WHERE) para poder añadir LEFT JOINs adicionales antes del filtro
			string baseFromNoWhere = $@"
				FROM [dbo].[{filesTable}] f
				JOIN [dbo].[{projectsTable}] p ON f.[{filesProjectFkColumn}] = p.[{projectPkColumn}]
				LEFT JOIN [dbo].[MetaExtractors] me_direct ON f.[Metadata_Id] = me_direct.[Id]
				LEFT JOIN [dbo].[MetaDatas] md ON f.[Metadata_Id] = md.[Id]
				LEFT JOIN [dbo].[MetaExtractors] me ON me.[FoundMetaData_Id] = md.[Id]";
			string whereClause = $"\n\t\t\tWHERE p.[{projectPkColumn}] = @ProjectId";

			string meUsersIdExpr = "ISNULL(me_direct.[FoundUsers_Id], me.[FoundUsers_Id])";
			string meEmailsIdExpr = "ISNULL(me_direct.[FoundEmails_Id], me.[FoundEmails_Id])";

			// 1) Fila base por fichero
			string baseFileSelect = $@"
				SELECT DISTINCT
					{ficheroCol} AS [Fichero],
					{urlCol} AS [URL],
					'' AS [Usuario],
					{ubicacionCol} AS [Carpeta],
					'' AS [Software],
					'' AS [Emails],
					'' AS [Clientes (equipos)]
				{baseFromNoWhere}{whereClause}";

			// 2) Usuarios (una fila por usuario)
			string usersSelect = enableUsers ? $@"
				SELECT DISTINCT
					{ficheroCol} AS [Fichero],
					{urlCol} AS [URL],
					ui.Name AS [Usuario],
					{ubicacionCol} AS [Carpeta],
					'' AS [Software],
					'' AS [Emails],
					'' AS [Clientes (equipos)]
				{baseFromNoWhere}
				LEFT JOIN [dbo].[Users] u ON u.[Id] = {meUsersIdExpr}
				LEFT JOIN [dbo].[{userItemsJoinTable}] ui ON ui.[Users_Id] = u.[Id]{whereClause}" : "";

			// 3) Emails (una fila por email)
			string emailsSelect = enableEmails ? $@"
				SELECT DISTINCT
					{ficheroCol} AS [Fichero],
					{urlCol} AS [URL],
					'' AS [Usuario],
					{ubicacionCol} AS [Carpeta],
					'' AS [Software],
					ei.Mail AS [Emails],
					'' AS [Clientes (equipos)]
				{baseFromNoWhere}
				LEFT JOIN [dbo].[Emails] e ON e.[Id] = {meEmailsIdExpr}
				LEFT JOIN [dbo].[{emailItemsJoinTable}] ei ON ei.[Emails_Id] = e.[Id]{whereClause}" : "";

			// 4) Software (una fila por software)
			string softwareSelect = enableApplications ? $@"
				SELECT DISTINCT
					{ficheroCol} AS [Fichero],
					{urlCol} AS [URL],
					'' AS [Usuario],
					{ubicacionCol} AS [Carpeta],
					ai.Name AS [Software],
					'' AS [Emails],
					'' AS [Clientes (equipos)]
				{baseFromNoWhere}
				LEFT JOIN [dbo].[{applicationItemsJoinTable}] ai ON ai.[Applications_Id] = md.[Applications_Id]{whereClause}" : "";

			// 5) Clientes (equipos) (una fila por equipo)
			string clientsSelect = enableComputers ? $@"
				SELECT DISTINCT
					{ficheroCol} AS [Fichero],
					{urlCol} AS [URL],
					'' AS [Usuario],
					{ubicacionCol} AS [Carpeta],
					'' AS [Software],
					'' AS [Emails],
					COALESCE(ci_id.name, ci_name.name) AS [Clientes (equipos)]
				{baseFromNoWhere}
				LEFT JOIN [dbo].[Users] u_c ON u_c.[Id] = {meUsersIdExpr}
				LEFT JOIN [dbo].[{userItemsJoinTable}] ui_c ON ui_c.[Users_Id] = u_c.[Id]
				LEFT JOIN [dbo].[{computersItemsTable}] ci_id ON ci_id.[IdProject] = f.[{filesProjectFkColumn}] AND ci_id.[type] = 0 AND (ci_id.[Users_Id] = {meUsersIdExpr} OR ci_id.[RemoteUsers_Id] = {meUsersIdExpr})
				LEFT JOIN [dbo].[{computersItemsTable}] ci_name ON ci_name.[IdProject] = f.[{filesProjectFkColumn}] AND ci_name.[type] = 0 AND UPPER(REPLACE(ci_name.[name],' ','')) COLLATE Latin1_General_CI_AI = UPPER('PC_'+REPLACE(ui_c.[Name],' ','')) COLLATE Latin1_General_CI_AI{whereClause}" : "";

			// 6) Folders (una fila por carpeta desde PathsItems)
			string foldersSelect = $@"
				SELECT DISTINCT
					{ficheroCol} AS [Fichero],
					{urlCol} AS [URL],
					'' AS [Usuario],
					pi.[Path] AS [Carpeta],
					'' AS [Software],
					'' AS [Emails],
					'' AS [Clientes (equipos)]
				{baseFromNoWhere}
				LEFT JOIN [dbo].[Paths] pth ON pth.[Id] = ISNULL(me_direct.[FoundPaths_Id], me.[FoundPaths_Id])
				LEFT JOIN [dbo].[PathsItems] pi ON pi.[Paths_Id] = pth.[Id]
				{whereClause} AND pi.[Path] IS NOT NULL";

			var parts = new System.Collections.Generic.List<string>();
			parts.Add(baseFileSelect);
			if (!string.IsNullOrEmpty(usersSelect)) parts.Add(usersSelect);
			if (!string.IsNullOrEmpty(emailsSelect)) parts.Add(emailsSelect);
			if (!string.IsNullOrEmpty(softwareSelect)) parts.Add(softwareSelect);
			if (!string.IsNullOrEmpty(clientsSelect)) parts.Add(clientsSelect);
			parts.Add(foldersSelect);
			query = string.Join("\nUNION ALL\n", parts) + "\nORDER BY 1,2";

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

        private static async Task<string> TryResolveFileNameFromUrlAsync(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return null;
            if (_urlToFileNameCache.TryGetValue(url, out var cached)) return cached;
            // HEAD primero
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "HEAD";
                req.Timeout = 5000;
                req.UserAgent = "FocaExcelExport/1.0";
                using (var resp = (HttpWebResponse)await req.GetResponseAsync())
                {
                    var name = GetFileNameFromContentDisposition(resp.Headers["Content-Disposition"]);
                    if (!string.IsNullOrWhiteSpace(name)) { _urlToFileNameCache[url] = name; return name; }
                }
            }
            catch { }

            // GET parcial como fallback
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                req.Timeout = 5000;
                req.UserAgent = "FocaExcelExport/1.0";
                req.AddRange(0, 0);
                using (var resp = (HttpWebResponse)await req.GetResponseAsync())
                {
                    var name = GetFileNameFromContentDisposition(resp.Headers["Content-Disposition"]);
                    if (!string.IsNullOrWhiteSpace(name)) { _urlToFileNameCache[url] = name; return name; }
                }
            }
            catch { }

            // Último segmento de URL
            try
            {
                var uri = new Uri(url, UriKind.RelativeOrAbsolute);
                var last = uri.IsAbsoluteUri ? uri.Segments[uri.Segments.Length - 1] : url.Substring(url.LastIndexOf('/') + 1);
                if (!string.IsNullOrWhiteSpace(last)) { _urlToFileNameCache[url] = last; return last; }
            }
            catch { }
            return null;
        }
    }
}