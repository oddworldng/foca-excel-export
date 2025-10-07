using ClosedXML.Excel;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using FocaExcelExport.Classes;

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

                // Count total records for progress
                var totalRecords = await GetRecordCountAsync(projectId, projectsTable, filesTable, metadataTable, 
                    projectPkColumn, filePkColumn, filesProjectFkColumn);
                
                int currentRecord = 0;
                progress?.Report(new ExportProgress 
                { 
                    CurrentRecord = currentRecord, 
                    TotalRecords = totalRecords, 
                    StatusMessage = "Initializing export..." 
                });

                // Build dynamic SQL query
                var query = BuildExportQuery(projectsTable, filesTable, metadataTable, 
                    projectPkColumn, filePkColumn, filesProjectFkColumn, fileNameColumn, urlColumn, 
                    userNameColumn, locationColumn, emailColumn, clientColumn, emailItemsTable, userItemsTable);

                // Create Excel file
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Exported Data");
                    
                    // Add headers
                    worksheet.Cell(1, 1).Value = "Fichero";
                    worksheet.Cell(1, 2).Value = "URL";
                    worksheet.Cell(1, 3).Value = "Usuario";
                    worksheet.Cell(1, 4).Value = "Ubicación";
                    worksheet.Cell(1, 5).Value = "Email";
                    worksheet.Cell(1, 6).Value = "Cliente";

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
                                    worksheet.Cell(row, 1).Value = reader[0]?.ToString() ?? "";
                                    worksheet.Cell(row, 2).Value = reader[1]?.ToString() ?? "";
                                    worksheet.Cell(row, 3).Value = reader[2]?.ToString() ?? "";
                                    worksheet.Cell(row, 4).Value = reader[3]?.ToString() ?? "";
                                    worksheet.Cell(row, 5).Value = reader[4]?.ToString() ?? "";
                                    worksheet.Cell(row, 6).Value = reader[5]?.ToString() ?? "";

                                    currentRecord++;
                                    row++;

                                    // Report progress
                                    if (progress != null && totalRecords > 0)
                                    {
                                        progress.Report(new ExportProgress
                                        {
                                            CurrentRecord = currentRecord,
                                            TotalRecords = totalRecords,
                                            StatusMessage = $"Processing record {currentRecord} of {totalRecords}..."
                                        });
                                    }
                                }
                            }
                        }
                    }

                    // Format header row
                    var headerRange = worksheet.Range(1, 1, 1, 6);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();

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
            string userNameColumn, string locationColumn, string emailColumn, string clientColumn, string emailItemsTable, string userItemsTable)
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
            string ficheroCol = "CASE WHEN f.[Path] IS NOT NULL AND LEN(f.[Path]) > 0 " +
                                "THEN RIGHT(f.[Path], CHARINDEX('\\\\', REVERSE(f.[Path]) + '\\\\') - 1) " +
                                "ELSE RIGHT(f.[URL], CHARINDEX('/', REVERSE(f.[URL]) + '/') - 1) END";
            // For URL, use the URL column
            string urlCol = "f.[URL]";
            // For Usuario (user), join through MetaExtractors -> Users -> UserItems
            string usuarioCol = "ISNULL(ui.Name, '')";
            // For Ubicación (location), use Path column
            string ubicacionCol = "f.[Path]";
            // For Email, join through MetaExtractors -> Emails -> EmailItems
            string emailCol = "ISNULL(ei.Mail, '')";
            // For Cliente (client), use Domain from Projects
            string clienteCol = "p.[Domain]";

            string query;
            var userItemsJoinTable = string.IsNullOrEmpty(userItemsTable) ? "UserItems" : userItemsTable;
            var emailItemsJoinTable = string.IsNullOrEmpty(emailItemsTable) ? "EmailItems" : emailItemsTable;
            
            // Handle case where metadata table is found
            if (!string.IsNullOrEmpty(metadataTable))
            {
                // Complex query joining files, projects, and metadata tables
                // FilesITems.Metadata_Id -> MetaExtractors.Id
                query = $@"
                    SELECT
                        {ficheroCol} AS [Fichero],
                        {urlCol} AS [URL],
                        {usuarioCol} AS [Usuario],
                        {ubicacionCol} AS [Ubicación],
                        {emailCol} AS [Email],
                        {clienteCol} AS [Cliente]
                    FROM [dbo].[{filesTable}] f
                    JOIN [dbo].[{projectsTable}] p ON f.[{filesProjectFkColumn}] = p.[{projectPkColumn}]
                    LEFT JOIN [dbo].[MetaExtractors] me ON f.[Metadata_Id] = me.[Id]
                    LEFT JOIN [dbo].[Users] u ON me.[FoundUsers_Id] = u.[Id]
                    LEFT JOIN [dbo].[{userItemsJoinTable}] ui ON ui.[Users_Id] = u.[Id]
                    LEFT JOIN [dbo].[Emails] e ON me.[FoundEmails_Id] = e.[Id]
                    LEFT JOIN [dbo].[{emailItemsJoinTable}] ei ON ei.[Emails_Id] = e.[Id]
                    WHERE p.[{projectPkColumn}] = @ProjectId
                    ORDER BY {ficheroCol}";
            }
            else
            {
                // Simplified query without metadata
                query = $@"
                    SELECT
                        {ficheroCol} AS [Fichero],
                        {urlCol} AS [URL],
                        '' AS [Usuario],
                        {ubicacionCol} AS [Ubicación],
                        '' AS [Email],
                        {clienteCol} AS [Cliente]
                    FROM [dbo].[{filesTable}] f
                    JOIN [dbo].[{projectsTable}] p ON f.[{filesProjectFkColumn}] = p.[{projectPkColumn}]
                    WHERE p.[{projectPkColumn}] = @ProjectId
                    ORDER BY {ficheroCol}";
            }

            return query;
        }
    }
}