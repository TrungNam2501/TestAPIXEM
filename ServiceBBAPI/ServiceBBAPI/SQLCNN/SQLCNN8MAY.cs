using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;

namespace ServiceBBAPI.SQLCNN
{

    public class SQLCNN8MAY
    {
        public static Dictionary<string, string> cnnstr = new Dictionary<string, string>
        {
            { "V-BB3701", @"Data Source=198.1.8.21;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3702", @"Data Source=198.1.8.22;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3703", @"Data Source=198.1.8.23;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3704", @"Data Source=198.1.8.24;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3705", @"Data Source=198.1.8.35;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3706", @"Data Source=198.1.8.36;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3707", @"Data Source=198.1.8.37;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "V-BB3708", @"Data Source=198.1.8.38;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "01", @"Data Source=198.1.8.21;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "02", @"Data Source=198.1.8.22;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "03", @"Data Source=198.1.8.23;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "04", @"Data Source=198.1.8.24;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "05", @"Data Source=198.1.8.35;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "06", @"Data Source=198.1.8.36;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "07", @"Data Source=198.1.8.37;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "08", @"Data Source=198.1.8.38;Initial Catalog=mfns;User ID=kendakv2;Password=kenda123" },
            { "33", @"Data Source=198.1.10.33;Initial Catalog=erp;User ID=kendakv2;Password=kenda123" },
            { "186", @"Data Source=198.1.9.186;Initial Catalog=InTem;User ID=kendakv2;Password=kenda123" },
            { "maytest", @"Data Source=198.1.10.133;Initial Catalog=LotBB;User ID=kenda;Password=kenda123" },
            { "34", @"Data Source=198.1.10.34;Initial Catalog=P8400;User ID=kendakv2;Password=kenda123" },
            { "V11", @"Data Source=198.1.8.16;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "V12", @"Data Source=198.1.8.17;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "V13", @"Data Source=198.1.8.15;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },
            { "V14", @"Data Source=198.1.8.18;Initial Catalog=CWSS_S7;User ID=kendakv2;Password=kenda123" },


        };
        public static DataTable ExecuteQuery(string Query,string may, object[] parameter = null)
        {
            cnnstr.TryGetValue(may, out string ConnectionString);

            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(Query, conn);
                    if (parameter != null)
                    {
                        string[] listPara = Query.Split(' ');
                        int i = 0;
                        foreach (string item in listPara)
                        {
                            if (item.Contains('?'))
                            {
                                cmd.Parameters.AddWithValue(item, parameter[i]);
                                i++;
                            }
                        }
                    }
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    return dt;
                }
                catch (Exception ex)
                {
                    return new DataTable();
                }
                finally
                {
                    if (conn.State != ConnectionState.Closed)
                        conn.Close();
                }
            }

        }

        public static bool ExecuteNonQuery(string query, string may, object[] parameter = null)
        {
            cnnstr.TryGetValue(may, out string ConnectionString);
        

            //string ConnectionString = "Data Source=198.1.10.34;Initial Catalog=erp;User ID=kendakv2;Password=kenda123";
            //string ConnectionString = @"Data Source = 198.1.1.95; Initial Catalog = JianDaMES; User ID = kendakv2; Password = kenda123";
            using (var conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    conn.Open();

                    SqlCommand cmd = new SqlCommand(query, conn);

                    if (parameter != null)
                    {
                        string[] listPara = query.Split(' ');
                        int i = 0;
                        foreach (string item in listPara)
                        {
                            if (item.Contains('?'))
                            {
                                cmd.Parameters.AddWithValue(item, parameter[i]);
                                i++;
                            }
                        }
                    }
                    int effectedRow = cmd.ExecuteNonQuery();
                    return effectedRow > 0;
                }
                catch (Exception ex)
                {
                    return false;
                }
                finally
                {
                    if (conn.State != ConnectionState.Closed)
                        conn.Close();
                }
            }
        }
    }
}