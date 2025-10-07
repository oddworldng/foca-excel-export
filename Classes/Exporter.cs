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
                return TotalRecords > 0 ? (int)((double)CurrentRecord / TotalRecords * 100) : 0; 
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

                // Find relevant columns
                var projectPkColumn = await schemaResolver.FindProjectIdColumnAsync(projectsTable);
                var filePkColumn = await schemaResolver.FindFileIdColumnAsync(filesTable);
                
                var fileNameColumn = await schemaResolver.FindFileNameColumnAsync(filesTable);
                var urlColumn = await schemaResolver.FindUrlColumnAsync(filesTable);
                
                var userNameColumn = await schemaResolver.FindUserNameColumnAsync(metadataTable);
                var locationColumn = await schemaResolver.FindLocationColumnAsync(metadataTable);
                var emailColumn = await schemaResolver.FindEmailColumnAsync(metadataTable);
                var clientColumn = await schemaResolver.FindClientColumnAsync(projectsTable);

                // Count total records for progress
                var totalRecords = await GetRecordCountAsync(projectId, projectsTable, filesTable, metadataTable, 
                    projectPkColumn, filePkColumn);
                
                int currentRecord = 0;
                progress?.Report(new ExportProgress 
                { 
                    CurrentRecord = currentRecord, 
                    TotalRecords = totalRecords, 
                    StatusMessage = "Initializing export..." 
                });

                // Build dynamic SQL query
                var query = BuildExportQuery(projectsTable, filesTable, metadataTable, 
                    projectPkColumn, filePkColumn, fileNameColumn, urlColumn, 
                    userNameColumn, locationColumn, emailColumn, clientColumn);

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
            string projectPkColumn, string filePkColumn)
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
                        JOIN [dbo].[{projectsTable}] p ON f.[{projectPkColumn}] = p.[{projectPkColumn}]
                        LEFT JOIN [dbo].[{metadataTable}] m ON m.[{filePkColumn}] = f.[{filePkColumn}]
                        WHERE p.[{projectPkColumn}] = @ProjectId";
                }
                else
                {
                    // Exclude metadata table
                    countQuery = $@"
                        SELECT COUNT(*)
                        FROM [dbo].[{filesTable}] f
                        JOIN [dbo].[{projectsTable}] p ON f.[{projectPkColumn}] = p.[{projectPkColumn}]
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
            string projectPkColumn, string filePkColumn, string fileNameColumn, string urlColumn,
            string userNameColumn, string locationColumn, string emailColumn, string clientColumn)
        {
            // Build the select clause with COALESCE for optional fields
            string ficheroCol = !string.IsNullOrEmpty(fileNameColumn) ? $"f.[{fileNameColumn}]" : "''";
            string urlCol = !string.IsNullOrEmpty(urlColumn) ? $"f.[{urlColumn}]" : "''";
            string usuarioCol = !string.IsNullOrEmpty(userNameColumn) ? $"COALESCE(m.[{userNameColumn}], '')" : "''";
            string ubicacionCol = !string.IsNullOrEmpty(locationColumn) ? $"COALESCE(m.[{locationColumn}], '')" : "''";
            string emailCol = !string.IsNullOrEmpty(emailColumn) ? $"COALESCE(m.[{emailColumn}], '')" : "''";
            string clienteCol = !string.IsNullOrEmpty(clientColumn) ? $"COALESCE(p.[{clientColumn}], '')" : "''";

            // Base query joining files and projects
            string query = $@"
                SELECT
                    {ficheroCol} AS [Fichero],
                    {urlCol} AS [URL],
                    {usuarioCol} AS [Usuario],
                    {ubicacionCol} AS [Ubicación],
                    {emailCol} AS [Email],
                    {clienteCol} AS [Cliente]
                FROM [dbo].[{filesTable}] f
                JOIN [dbo].[{projectsTable}] p ON f.[{projectPkColumn}] = p.[{projectPkColumn}]";

            // Add metadata table if it exists
            if (!string.IsNullOrEmpty(metadataTable))
            {
                query += $" LEFT JOIN [dbo].[{metadataTable}] m ON m.[{filePkColumn}] = f.[{filePkColumn}]";
            }

            query += $" WHERE p.[{projectPkColumn}] = @ProjectId ORDER BY {ficheroCol}";

            return query;
        }
    }
}