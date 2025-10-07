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
        /// Find the projects table based on common column names
        /// </summary>
        public async Task<string> FindProjectsTableAsync()
        {
            var tables = await GetTablesAsync();
            
            foreach (var table in tables)
            {
                var columns = await GetColumnsAsync(table);
                
                // Look for project-related columns
                if (columns.Any(c => c.ToLower().Contains("project") && c.ToLower().Contains("id")) ||
                    columns.Any(c => c.ToLower().Contains("name")) ||
                    columns.Any(c => c.ToLower().Contains("title")) ||
                    columns.Any(c => c.ToLower().Contains("project")))
                {
                    // Additional check for project characteristics
                    if (columns.Contains("Id") || columns.Any(c => c.ToLower().Contains("projectid")))
                    {
                        return table;
                    }
                }
            }
            
            // If no specific project table found, look for common names
            var possibleProjectTables = new[] { "Projects", "Project", "TblProjects", "ProjectInfo", "tblProjects" };
            foreach (var possibleTable in possibleProjectTables)
            {
                if (tables.Contains(possibleTable))
                {
                    return possibleTable;
                }
            }
            
            // Return the first table that has both an ID and a name field
            foreach (var table in tables)
            {
                var columns = await GetColumnsAsync(table);
                if (columns.Contains("Id") && (columns.Any(c => c.ToLower().Contains("name")) || columns.Any(c => c.ToLower().Contains("title"))))
                {
                    return table;
                }
            }
            
            return tables.FirstOrDefault(); // Fallback
        }

        /// <summary>
        /// Find the files table based on common column names
        /// </summary>
        public async Task<string> FindFilesTableAsync()
        {
            var tables = await GetTablesAsync();
            
            foreach (var table in tables)
            {
                var columns = await GetColumnsAsync(table);
                
                // Look for file-related columns
                if (columns.Any(c => c.ToLower().Contains("file") && c.ToLower().Contains("id")) ||
                    columns.Any(c => c.ToLower().Contains("filename")) ||
                    columns.Any(c => c.ToLower().Contains("filepath")) ||
                    columns.Any(c => c.ToLower().Contains("url")) ||
                    columns.Any(c => c.ToLower().Contains("file")))
                {
                    // Additional check for file characteristics
                    if (columns.Contains("Id") || columns.Any(c => c.ToLower().Contains("fileid")))
                    {
                        return table;
                    }
                }
            }
            
            // If no specific file table found, look for common names
            var possibleFileTables = new[] { "Files", "File", "TblFiles", "FileInfo", "Documents", "tblFiles" };
            foreach (var possibleTable in possibleFileTables)
            {
                if (tables.Contains(possibleTable))
                {
                    return possibleTable;
                }
            }
            
            // Return the first table that has file characteristics
            foreach (var table in tables)
            {
                var columns = await GetColumnsAsync(table);
                if (columns.Any(c => c.ToLower().Contains("file")) || columns.Any(c => c.ToLower().Contains("url")))
                {
                    return table;
                }
            }
            
            return tables.FirstOrDefault(); // Fallback
        }

        /// <summary>
        /// Find the metadata table based on common column names
        /// </summary>
        public async Task<string> FindMetadataTableAsync()
        {
            var tables = await GetTablesAsync();
            
            foreach (var table in tables)
            {
                var columns = await GetColumnsAsync(table);
                
                // Look for metadata-related columns (user, email, location, etc.)
                var hasUserRelated = columns.Any(c => c.ToLower().Contains("user") || c.ToLower().Contains("name"));
                var hasEmail = columns.Any(c => c.ToLower().Contains("email"));
                var hasLocation = columns.Any(c => c.ToLower().Contains("location") || c.ToLower().Contains("path"));
                
                if (hasUserRelated && (hasEmail || hasLocation))
                {
                    return table;
                }
            }
            
            // If no specific metadata table found, look for common names
            var possibleMetadataTables = new[] { "Metadata", "FileMetadata", "UserMetadata", "DocumentMetadata", "UserInfo", "TblMetadata" };
            foreach (var possibleTable in possibleMetadataTables)
            {
                if (tables.Contains(possibleTable))
                {
                    return possibleTable;
                }
            }
            
            // Return the first table with user-related data
            foreach (var table in tables)
            {
                var columns = await GetColumnsAsync(table);
                if (columns.Any(c => c.ToLower().Contains("user")) || 
                    columns.Any(c => c.ToLower().Contains("email")) || 
                    columns.Any(c => c.ToLower().Contains("location")))
                {
                    return table;
                }
            }
            
            return null; // It's ok if no metadata table is found
        }

        /// <summary>
        /// Find the column name for project ID in a table
        /// </summary>
        public async Task<string> FindProjectIdColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleProjectIdColumns = new[] { "ProjectId", "Project_ID", "Project_Id", "projectId", "projectid", "Projectid" };
            foreach (var col in possibleProjectIdColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column containing both project and id
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("project") && col.ToLower().Contains("id"))
                {
                    return col;
                }
            }
            
            return columns.FirstOrDefault(c => c.ToLower().Contains("project")); // Fallback
        }

        /// <summary>
        /// Find the column name for file ID in a table
        /// </summary>
        public async Task<string> FindFileIdColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleFileIdColumns = new[] { "FileId", "File_ID", "File_Id", "fileId", "fileid", "Fileid" };
            foreach (var col in possibleFileIdColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column containing both file and id
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("file") && col.ToLower().Contains("id"))
                {
                    return col;
                }
            }
            
            return columns.FirstOrDefault(c => c.ToLower().Contains("file")); // Fallback
        }

        /// <summary>
        /// Find the column name for file name in a table
        /// </summary>
        public async Task<string> FindFileNameColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleFileNameColumns = new[] { "FileName", "File_Name", "filename", "Name", "name", "Title", "title", "RealName" };
            foreach (var col in possibleFileNameColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column that might contain file information
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("file") && !col.ToLower().Contains("id"))
                {
                    return col;
                }
                if (col.ToLower().Contains("name") && !col.ToLower().Contains("user"))
                {
                    return col;
                }
            }
            
            return columns.FirstOrDefault(); // Fallback
        }

        /// <summary>
        /// Find the column name for URL in a table
        /// </summary>
        public async Task<string> FindUrlColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleUrlColumns = new[] { "Url", "URL", "url", "FullUrl", "Full_URL", "Path", "URLPath", "FileUrl", "DocumentUrl" };
            foreach (var col in possibleUrlColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column that might contain URL information
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("url") || col.ToLower().Contains("path"))
                {
                    return col;
                }
            }
            
            return null; // May not exist in this table
        }

        /// <summary>
        /// Find the column name for user name in a table
        /// </summary>
        public async Task<string> FindUserNameColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleUserColumns = new[] { 
                "UserName", "User_Name", "username", "User", "user", "Name", "name", 
                "FullName", "First_Name", "Last_Name", "Author", "Owner", "CreatedBy", "Creator" 
            };
            foreach (var col in possibleUserColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column that might contain user information
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("user") || col.ToLower().Contains("name") || col.ToLower().Contains("author"))
                {
                    return col;
                }
            }
            
            return null; // May not exist in this table
        }

        /// <summary>
        /// Find the column name for location/path in a table
        /// </summary>
        public async Task<string> FindLocationColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleLocationColumns = new[] { 
                "Location", "location", "Path", "path", "FilePath", "File_Path", "Directory", 
                "Dir", "NetworkPath", "Network_Path", "LocationPath", "Location_Path" 
            };
            foreach (var col in possibleLocationColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column that might contain location information
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("location") || col.ToLower().Contains("path") || 
                    col.ToLower().Contains("dir") || col.ToLower().Contains("folder"))
                {
                    return col;
                }
            }
            
            return null; // May not exist in this table
        }

        /// <summary>
        /// Find the column name for email in a table
        /// </summary>
        public async Task<string> FindEmailColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleEmailColumns = new[] { "Email", "email", "EmailAddr", "EmailAddress", "EMail", "Email_Address" };
            foreach (var col in possibleEmailColumns)
            {
                if (columns.Contains(col))
                {
                    return col;
                }
            }
            
            // Look for any column that might contain email information
            foreach (var col in columns)
            {
                if (col.ToLower().Contains("email"))
                {
                    return col;
                }
            }
            
            return null; // May not exist in this table
        }

        /// <summary>
        /// Find the column name for client name in a table
        /// </summary>
        public async Task<string> FindClientColumnAsync(string tableName)
        {
            var columns = await GetColumnsAsync(tableName);
            
            var possibleClientColumns = new[] { 
                "Client", "client", "ClientName", "Client_Name", "Company", "Organization", 
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
    }
}