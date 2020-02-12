using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace MyCommands
{
    public class MyCommandsDAL
    {
        public string getVersionPath(string versionNumber)
        {
            SqlConnection sqlConnection1 = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            cmd.CommandText = $"SELECT VersionPath FROM tblVersion where versionNumber='{versionNumber}'";
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();

            reader = cmd.ExecuteReader();

            string results = "";

            // Data is accessible through the DataReader object here..
            if (reader.Read())
            {
                results = reader["VersionPath"].ToString();
            }

            sqlConnection1.Close();
            return results;
        }

        public List<int> getAllVersionNumbers()
        {
            SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
            SqlCommand command = new SqlCommand("select versionNumber from tblVersion", con);
            List<int> ints = new List<int>();
            con.Open();

            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    ints.Add(reader.GetInt32(0)); // provided that first (0-index) column is int which you are looking for
                }
            }

            con.Close();
            return ints;
        }

        public void SaveExecutedCommand(string commandText, string commandNumber)
        {
            SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
            SqlCommand cmd;

            cmd = new SqlCommand("insert into tblCommands(commandNumber,CommandText, LastModifiedTimestamp) values(@commandNumber,@CommandText, @LastModifiedTimestamp)", con);
            con.Open();
            cmd.Parameters.AddWithValue("@commandNumber", commandNumber);
            cmd.Parameters.AddWithValue("@CommandText", commandText);
            cmd.Parameters.AddWithValue("@LastModifiedTimestamp", DateTime.Now);
            cmd.ExecuteNonQuery();
            con.Close();
        }

        public bool IsCommandAlreadyExecuted(string commandNumber)
        {
            SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
            SqlCommand command = new SqlCommand($"select 1 from tblCommands where CommandNumber = '{commandNumber}'", con);
            con.Open();
            bool result = false;

            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    result = true;
                }
            }

            con.Close();
            return result;
        }

    }
}
