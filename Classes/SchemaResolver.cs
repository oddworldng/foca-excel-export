using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace FocaExcelExport.Classes
{
    public class SchemaResolver
    {
        private readonly string _connectionString;

        public SchemaResolver(string connectionString)
        {
            _connectionString = connectionString;
        }

        /// <summary>
        /// Discover all tables in the database
        /// </summary>
        public async Task<List<string>> GetTablesAsync()
        {
            var tables = new List<string>();
            
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                
                var command = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE'", connection);
                var reader = await command.ExecuteReaderAsync();
                
                while (await reader.ReadAsync())
                {
                    tables.Add(reader.GetString(0));
                }
            }
            
            return tables;
        }

        /// <summary>
        /// Check if a table exists by exact name
        /// </summary>
        public async Task<bool> TableExistsAsync(string tableName)
        {
            var tables = await GetTablesAsync();
            return tables.Contains(tableName);
        }

        /// <summary>
        /// Get all columns for a specific table
        /// </summary>
        public async Task<List<string>> GetColumnsAsync(string tableName)
        {
            var columns = new List<string>();
            
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                
                var command = new SqlCommand($"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME=@TableName", connection);
                command.Parameters.AddWithValue("@TableName", tableName);
                var reader = await command.ExecuteReaderAsync();
                
                while (await reader.ReadAsync())
                {
                    columns.Add(reader.GetString(0));
                }
            }
            
            return columns;
        }

        /// <summary>
        /// Find the projects table based on FOCA Entity Framework structure
        /// From migrations: Table "Projects" with columns Id and ProjectName
        /// </summary>
        public async Task<string> FindProjectsTableAsync()
        {
            var tables = await GetTablesAsync();
            
            // Based on migration file, FOCA uses exact table name "Projects"
            if (tables.Contains("Projects"))
            {
                return "Projects";
            }
            
            return tables.FirstOrDefault(); // Fallback
        }

        /// <summary>
        /// Find the files table based on FOCA Entity Framework structure
        /// From migrations: Table "FilesITems" (note the 's' at the end) with columns Id, IdProject, URL, Path, etc.
        /// </summary>
        public async Task<string> FindFilesTableAsync()
        {
            var tables = await GetTablesAsync();
            
            // Based on migration file, FOCA uses exact table name "FilesITems" (note the 's' at the end)
            if (tables.Contains("FilesITems"))
            {
                return "FilesITems";
            }
            
            return tables.FirstOrDefault(); // Fallback
        }

        /// <summary>
        /// Find the likely foreign key column in files table that references projects (e.g., IdProject, ProjectId)
        /// </summary>
        public async Task<string> FindFilesProjectFkColumnAsync(string filesTable)
        {
            var columns = await GetColumnsAsync(filesTable);

            // Prefer exact common names first
            var preferred = new[] { "IdProject", "ProjectId", "Projects_Id" };
            foreach (var name in preferred)
            {
                if (columns.Contains(name)) return name;
            }

            // Heuristic: contains both "project" and "id"
            var heuristic = columns.FirstOrDefault(c => c.IndexOf("project", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                                        c.IndexOf("id", StringComparison.OrdinalIgnoreCase) >= 0);
            if (!string.IsNullOrEmpty(heuristic)) return heuristic;

            // Fallback to first column that ends with Id and is not the primary key Id
            var alt = columns.FirstOrDefault(c => !string.Equals(c, "Id", StringComparison.OrdinalIgnoreCase) &&
                                                  c.EndsWith("Id", StringComparison.OrdinalIgnoreCase));
            return alt ?? "IdProject"; // conservative default
        }

        /// <summary>
        /// Find the metadata table based on FOCA Entity Framework structure
        /// From migrations: Table "MetaExtractors" with relationships to FoundUsers_Id, FoundEmails_Id, etc.
        /// </summary>
        public async Task<string> FindMetadataTableAsync()
        {
            var tables = await GetTablesAsync();
            
            // Based on migration file, FOCA uses exact table name "MetaExtractors"
            if (tables.Contains("MetaExtractors"))
            {
                return "MetaExtractors";
            }
            
            return null; // It's ok if no metadata table is found
        }

        /// <summary>
        /// Find EmailItems-like table (EmailItems or EmailsItems)
        /// </summary>
        public async Task<string> FindEmailItemsTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("EmailItems")) return "EmailItems";
            if (tables.Contains("EmailsItems")) return "EmailsItems";

            // Heuristic by columns
            foreach (var t in tables)
            {
                if (t.IndexOf("Email", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    t.IndexOf("Item", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var cols = await GetColumnsAsync(t);
                    if (cols.Contains("Mail")) return t;
                }
            }
            return null;
        }

        /// <summary>
        /// Find UserItems-like table (UserItems or UsersItems)
        /// </summary>
        public async Task<string> FindUserItemsTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("UserItems")) return "UserItems";
            if (tables.Contains("UsersItems")) return "UsersItems";

            foreach (var t in tables)
            {
                if (t.IndexOf("User", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    t.IndexOf("Item", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var cols = await GetColumnsAsync(t);
                    if (cols.Contains("Name")) return t;
                }
            }
            return null;
        }

        /// <summary>
        /// Find the column name for project ID in a table
        /// Based on FOCA Entity Framework, this should be "Id" (capitalized)
        /// </summary>
        public async Task<string> FindProjectIdColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on FOCA EF structure, the ID column is "Id" (capitalized)
            if (columns.Contains("Id"))
            {
                return "Id";
            }
            
            return "Id"; // Default based on Entity Framework conventions
        }

        /// <summary>
        /// Find the column name for file ID in a table
        /// </summary>
        public async Task<string> FindFileIdColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on FOCA Entity Framework, the ID column is "Id" (standard EF convention)
            if (columns.Contains("Id"))
            {
                return "Id";
            }
            
            return "Id"; // Default based on Entity Framework conventions
        }

        /// <summary>
        /// Find the column name for project name in a table
        /// Based on FOCA Project entity, this should be "ProjectName"
        /// </summary>
        public async Task<string> FindFileNameColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on FOCA Project entity structure, the project name column is "ProjectName"
            if (tableName == "Projects" && columns.Contains("ProjectName"))
            {
                return "ProjectName";
            }
            
            // For file names in files table, the URL column contains the file reference
            if (tableName == "FilesITems" && columns.Contains("URL"))
            {
                return "URL";
            }
            
            return "ProjectName"; // Default for projects table
        }

        /// <summary>
        /// Find the column name for URL in a table
        /// Based on FOCA FilesITems entity, this should be "URL"
        /// </summary>
        public async Task<string> FindUrlColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on FOCA FilesITems entity structure from migrations, the URL column is "URL"
            if (columns.Contains("URL"))
            {
                return "URL";
            }
            
            return "URL"; // Default based on FOCA structure
        }

        /// <summary>
        /// Find the column name for user name in a table
        /// Based on FOCA structure, users are in UserItems table with "Name" column
        /// </summary>
        public async Task<string> FindUserNameColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on migrations, UserItems table has "Name" column
            if (columns.Contains("Name"))
            {
                return "Name";
            }
            
            return "Name"; // Default based on FOCA UserItems structure
        }

        /// <summary>
        /// Find the column name for location/path in a table
        /// Based on FOCA FilesItem entity, this should be "Path"
        /// </summary>
        public async Task<string> FindLocationColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on FOCA FilesITems entity structure from migrations, the path column is "Path"
            if (columns.Contains("Path"))
            {
                return "Path";
            }
            
            return "Path"; // Default based on FOCA structure
        }

        /// <summary>
        /// Find the column name for email in a table
        /// Based on FOCA structure, emails are in EmailsItems table with "Mail" column
        /// </summary>
        public async Task<string> FindEmailColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            // Based on migrations, EmailsItems table has "Mail" column
            if (columns.Contains("Mail"))
            {
                return "Mail";
            }
            
            return "Mail"; // Default based on FOCA EmailsItems structure
        }

        /// <summary>
        /// Find the column name for client name in a table
        /// </summary>
        public async Task<string> FindClientColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleClientColumns = new[] { 
                "Domain", "Client", "client", "ClientName", "Client_Name", "Company", "Organization", 
                "Team", "TeamName", "OrganizationName", "Customer", "CustomerName", "Cliente" 
            };
            foreach (var col in possibleClientColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column that might contain client information
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("client") || col.ToLower().Contains("company") || 
                    col.ToLower().Contains("organization") || col.ToLower().Contains("team"))
                {
                    return col;
                }
            }
            
            return null; // May not exist in this table
        }

        // Additional helpers to discover Applications/Servers tables used to build Software/Equipos columns
        public async Task<string> FindApplicationsTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("Applications")) return "Applications";
            // Fallback by heuristic
            return tables.FirstOrDefault(t => t.IndexOf("Application", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public async Task<string> FindApplicationItemsTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("ApplicationItems")) return "ApplicationItems";
            if (tables.Contains("ApplicationsItems")) return "ApplicationsItems";
            foreach (var t in tables)
            {
                if (t.IndexOf("Application", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    t.IndexOf("Item", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var cols = await GetColumnsAsync(t);
                    if (cols.Contains("Name")) return t;
                }
            }
            return null;
        }

        public async Task<string> FindServersTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("Servers")) return "Servers";
            return tables.FirstOrDefault(t => t.IndexOf("Server", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public async Task<string> FindServerItemsTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("ServerItems")) return "ServerItems";
            if (tables.Contains("ServersItems")) return "ServersItems";
            foreach (var t in tables)
            {
                if (t.IndexOf("Server", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    t.IndexOf("Item", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var cols = await GetColumnsAsync(t);
                    if (cols.Contains("Name")) return t;
                }
            }
            return null;
        }

        public async Task<string> FindComputersItemsTableAsync()
        {
            var tables = await GetTablesAsync();
            if (tables.Contains("ComputersItems")) return "ComputersItems";
            // Heurística
            foreach (var t in tables)
            {
                if (t.IndexOf("Computer", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    t.EndsWith("Items", StringComparison.OrdinalIgnoreCase))
                {
                    var cols = await GetColumnsAsync(t);
                    if (cols.Contains("IdProject") && cols.Contains("name")) return t;
                }
            }
            return null;
        }
    }
}