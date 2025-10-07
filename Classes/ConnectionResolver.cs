using System;
using System.Configuration;
using System.Linq;

namespace FocaExcelExport.Classes
{
    public class ConnectionResolver
    {
        /// <summary>
        /// Finds the FOCA database connection string from the current AppDomain
        /// </summary>
        /// <returns>The connection string for the FOCA database</returns>
        public static string GetFocaConnectionString()
        {
            try
            {
                // Look for connection strings in the current AppDomain
                foreach (ConnectionStringSettings connectionString in ConfigurationManager.ConnectionStrings)
                {
                    // Look for a connection string that contains SQL Server specific elements
                    string cs = connectionString.ConnectionString.ToLower();
                    if (cs.Contains("data source") && cs.Contains("initial catalog"))
                    {
                        return connectionString.ConnectionString;
                    }
                }
                
                // If we couldn't find a connection string with the expected pattern, 
                // try to get the first non-default connection string
                var nonDefaultConnection = ConfigurationManager.ConnectionStrings
                    .Cast<ConnectionStringSettings>()
                    .FirstOrDefault(cs => !string.IsNullOrEmpty(cs.ConnectionString) && 
                                         !cs.ConnectionString.Contains("LocalSqlServer") && 
                                         !cs.ConnectionString.Contains("DefaultConnection"));
                
                return nonDefaultConnection?.ConnectionString;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Could not retrieve FOCA database connection string", ex);
            }
        }
    }
}