using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using Oracle.ManagedDataAccess.Client;

namespace APITest
{
    public class APIDatabase
    {
        public DataSet ConsultarSqlOra(string Sentencia, string user, string database)
        {
            DataSet dataSet = new DataSet();
            DataTable dTable = new DataTable("dTable");
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dataSet;
            }
            else if (database == "ORA")
            {
                OracleConnection cnn = new OracleConnection(ConnectionString2);
                cnn.Open();
                OracleDataAdapter adapter = new OracleDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dataSet;
            }
            return null;
        }

        public void UpdateDeleteInsert(string Sentencia, string Database, string User)
        {
            string ConnectionString = ConfigurationManager.ConnectionStrings[User].ConnectionString;
            switch (Database.ToUpper())
            {
                case "SQL":

                    SqlConnection sqlConnection = new SqlConnection(ConnectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    SqlTransaction sqlTransaction;
                    sqlTransaction = sqlConnection.BeginTransaction();
                    sqlCommand.Connection = sqlConnection;
                    sqlCommand.Transaction = sqlTransaction;
                    try
                    {
                        sqlCommand.CommandText = Sentencia;
                        sqlCommand.ExecuteNonQuery();
                        sqlTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        sqlTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    sqlConnection.Close();
                    break;

                case "ORA":

                    OracleConnection oracleConnection = new OracleConnection(ConnectionString);
                    oracleConnection.Open();
                    OracleCommand oracleCommand = oracleConnection.CreateCommand();
                    OracleTransaction oracleTransaction;
                    oracleTransaction = oracleConnection.BeginTransaction();
                    oracleCommand.Connection = oracleConnection;
                    oracleCommand.Transaction = oracleTransaction;
                    try
                    {
                        oracleCommand.CommandText = Sentencia;
                        oracleCommand.ExecuteNonQuery();
                        oracleTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        oracleTransaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    oracleConnection.Close();
                    break;
                default:
                    break;
            }



        }

        public DataTable Select(string Sentencia, string user, string database)
        {
            DataSet dataSet = new DataSet();
            DataTable dTable = new DataTable("dTable");
            string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;

            if (database == "SQL")
            {
                SqlConnection cnn = new SqlConnection(ConnectionString2);
                cnn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dTable;
            }
            else if (database == "ORA")
            {
                OracleConnection cnn = new OracleConnection(ConnectionString2);
                cnn.Open();
                OracleDataAdapter adapter = new OracleDataAdapter(Sentencia, cnn);
                adapter.Fill(dTable);
                dataSet.Tables.Add(dTable);
                cnn.Close();
                return dTable;
            }
            return null;

        }
    }
}
