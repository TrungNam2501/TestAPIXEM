using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace ServiceBBAPI.SQLCNN
{
    public class DatabaseHelper
    {
        private readonly string _connectionString;

        public DatabaseHelper(string connectionString)
        {
            _connectionString = connectionString;
        }

        public static DatabaseHelper FromConfigName(string configName)
        {
            string connStr = ConfigurationManager.ConnectionStrings[configName]?.ConnectionString;
            if (string.IsNullOrEmpty(connStr))
                throw new InvalidOperationException($"Connection string '{configName}' not found in configuration.");
            return new DatabaseHelper(connStr);
        }

        public DataTable ExecuteQuery(string query, params SqlParameter[] parameters)
        {
            using (var conn = new SqlConnection(_connectionString))
            {
                try
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 30;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        using (var adapter = new SqlDataAdapter(cmd))
                        {
                            var dt = new DataTable();
                            adapter.Fill(dt);
                            return dt;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DB Error] ExecuteQuery: {ex.Message}");
                    return new DataTable();
                }
            }
        }

        public bool ExecuteNonQuery(string query, params SqlParameter[] parameters)
        {
            using (var conn = new SqlConnection(_connectionString))
            {
                try
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 30;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        int affectedRows = cmd.ExecuteNonQuery();
                        return affectedRows > 0;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DB Error] ExecuteNonQuery: {ex.Message}");
                    return false;
                }
            }
        }

        public object ExecuteScalar(string query, params SqlParameter[] parameters)
        {
            using (var conn = new SqlConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = new SqlCommand(query, conn))
                {
                    cmd.CommandTimeout = 30;
                    if (parameters != null)
                    {
                        cmd.Parameters.AddRange(parameters);
                    }
                    return cmd.ExecuteScalar();
                }
            }
        }

        public bool CheckConnection()
        {
            using (var conn = new SqlConnection(_connectionString))
            {
                try
                {
                    conn.Open();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
    }

    public static class DatabaseConnections
    {
        private static readonly Dictionary<string, string> MachineConnectionMap = new Dictionary<string, string>
        {
            { "V-BB3701", "MFNS_8_21" }, { "01", "MFNS_8_21" },
            { "V-BB3702", "MFNS_8_22" }, { "02", "MFNS_8_22" },
            { "V-BB3703", "MFNS_8_23" }, { "03", "MFNS_8_23" },
            { "V-BB3704", "MFNS_8_24" }, { "04", "MFNS_8_24" },
            { "V-BB3705", "MFNS_8_35" }, { "05", "MFNS_8_35" },
            { "V-BB3706", "MFNS_8_36" }, { "06", "MFNS_8_36" },
            { "V-BB3707", "MFNS_8_37" }, { "07", "MFNS_8_37" },
            { "V-BB3708", "MFNS_8_38" }, { "08", "MFNS_8_38" },
            { "33", "ERP_10_33" },
            { "186", "InTem_9_186" },
            { "maytest", "LotBB_10_133" },
            { "34", "P8400_10_34" },
            { "V11", "CWSS_8_16" },
            { "V12", "CWSS_8_17" },
            { "V13", "CWSS_8_15" },
            { "V14", "CWSS_8_18" },
        };

        public static DatabaseHelper GetErpDb()
        {
            return DatabaseHelper.FromConfigName("ERP_10_33");
        }

        public static DatabaseHelper GetErp34Db()
        {
            return DatabaseHelper.FromConfigName("ERP_10_34");
        }

        public static DatabaseHelper GetInTemDb()
        {
            return DatabaseHelper.FromConfigName("InTem_9_186");
        }

        public static DatabaseHelper GetMachineDb(string machineKey)
        {
            if (MachineConnectionMap.TryGetValue(machineKey, out string configName))
            {
                return DatabaseHelper.FromConfigName(configName);
            }
            throw new ArgumentException($"Unknown machine key: {machineKey}");
        }
    }
}
