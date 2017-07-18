using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace MailMerger
{
    public class DatabaseInfo
    {

        public string database_name { get; set; }
        public string database_connection_string { get; set; }
        public string merge_working_path { get; set; }
        public string zip_path { get; set; }

        public static string dbaseConnection
        {
            get
            {
                string _connectionString = ConfigurationManager.ConnectionStrings["dbaseConnection"].ConnectionString;
                return _connectionString;
            }

        }
      

        public DatabaseInfo dbase(string pmdatabase)
        {
            DatabaseInfo dbase = new DatabaseInfo();

            try
            {
            DataSet result = new DataSet();// = string.Empty;

            SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["dbaseConnection"].ToString());
            sqlConnection.Open();
            SqlCommand select = new SqlCommand("getDatabaseDetails", sqlConnection);
            select.CommandType = CommandType.StoredProcedure;
            select.Connection = sqlConnection;
            SqlDataAdapter dataAdapter = new SqlDataAdapter(select);
            select.Parameters.Add("@pmdbase", SqlDbType.VarChar).Value = pmdatabase;

            dataAdapter.Fill(result);

            dbase.database_name = result != null ? result.Tables[0].Rows[0]["database_name"].ToString() : null;
            dbase.database_connection_string = result != null ? result.Tables[0].Rows[0]["database_connection_string"].ToString() : null;
            dbase.merge_working_path = result != null ? result.Tables[0].Rows[0]["merge_working_path"].ToString() : null;
            dbase.zip_path = result != null ? result.Tables[0].Rows[0]["zip_path"].ToString() : null;
            }
            catch (SqlException e)
            {
                MailMerge.WriteError("Error Connecting to database - " + e);
            }
            catch (Exception e)
            {
                MailMerge.WriteError("Error - " + e);
            }
      


            return dbase;
        }


        public static DataTable GetCVSValues(string scheme, string reportID, DatabaseInfo dbase)
        {
            DataSet result = new DataSet();// = string.Empty;

            SqlConnection sqlConnection = new SqlConnection(dbase.database_connection_string);
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            SqlCommand select = new SqlCommand("sp_Get_TemplateFieldsAsCSV");
            @select.CommandType = CommandType.StoredProcedure;
            @select.Connection = sqlConnection;
            dataAdapter.SelectCommand = @select;

            @select.Parameters.AddWithValue("@Scheme", scheme);
            @select.Parameters.AddWithValue("@ReportID", reportID);

            dataAdapter.Fill(result);

            return result.Tables[0];
        }

        public static void AddtoMailMergeQueue(int status, string request_string, string pid_id, string message, DatabaseInfo dbase, int total_records)
        {

            DataSet result = new DataSet();// = string.Empty;
            SqlConnection sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["dbaseConnection"].ToString());
            sqlConnection.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            SqlCommand select = new SqlCommand("AddtoMailMergeQueue");
            @select.CommandType = CommandType.StoredProcedure;
            @select.Connection = sqlConnection;
            dataAdapter.SelectCommand = @select;

            @select.Parameters.AddWithValue("@status", status);
            @select.Parameters.AddWithValue("@request_string", request_string);
            @select.Parameters.AddWithValue("@pid_id", pid_id);
            @select.Parameters.AddWithValue("@message", message);
            @select.Parameters.AddWithValue("@database_name", dbase.database_name);
            @select.Parameters.AddWithValue("@total_records", total_records);
            
            select.ExecuteNonQuery();
       }

        public static string[] GetSchemeReportID(string filePath, DatabaseInfo dbase)
        {
            string[] info = { "-", "-" };
            DataSet result = new DataSet();// = string.Empty;

            SqlConnection sqlConnection = new SqlConnection(dbase.database_connection_string);
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            SqlCommand select = new SqlCommand("sp_Get_SchemeAndReportID");
            @select.CommandType = CommandType.StoredProcedure;
            @select.Connection = sqlConnection;
            dataAdapter.SelectCommand = @select;

            @select.Parameters.AddWithValue("@filePath", filePath);

            dataAdapter.Fill(result);

            foreach (DataRow dr in result.Tables[0].Rows)
            {
                info[0] = dr[0].ToString();
                info[1] = dr[1].ToString();
            }

            return info;

        }

       
    }
}