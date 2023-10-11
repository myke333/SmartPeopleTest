using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.OracleClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;   
using System.Configuration;

namespace Web_Kactus_Test_V2
{
    class SqlAdapter
    {
        private static readonly string ConnectionString = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;

        static public DataSet SelectExecutionOrder(string p, string plan = null, string suite = null, string CountDes = null)
        {
            try
            {
                DataSet Data = new DataSet();
                SqlConnection cnn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("SelecOrderExecution_SelfService", cnn);

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Suite", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Plan", SqlDbType.Char, 50);
                cmd.Parameters.Add("@CountDes", SqlDbType.Char, 50);
                cmd.Parameters["@p"].Value = p;
                cmd.Parameters["@Suite"].Value = suite;
                cmd.Parameters["@Plan"].Value = plan;
                cmd.Parameters["@CountDes"].Value = CountDes;


                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);

                Data.Tables.Add(dTable);
                return Data;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }

        static public DataSet SelectExecutionOrderV2(string p, string plan = null, string suite = null, string CountDes = null, string CaseID = null)
        {
            try
            {
                DataSet Data = new DataSet();
                SqlConnection cnn = new SqlConnection(ConnectionString);
                SqlCommand cmd = new SqlCommand("SelecOrderExecutionV2_SelfService", cnn);

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@p", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Suite", SqlDbType.Char, 50);
                cmd.Parameters.Add("@Plan", SqlDbType.Char, 50);
                cmd.Parameters.Add("@CaseID", SqlDbType.Char, 50);
                cmd.Parameters.Add("@CountDes", SqlDbType.Char, 50);
                cmd.Parameters["@p"].Value = p;
                cmd.Parameters["@Suite"].Value = suite;
                cmd.Parameters["@Plan"].Value = plan;
                cmd.Parameters["@CaseID"].Value = CaseID;
                cmd.Parameters["@CountDes"].Value = CountDes;


                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);

                Data.Tables.Add(dTable);
                return Data;
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }

        static public DataSet SelectOrderExecution(string Parameter, string Table, string plan = null, string suite = null, string CaseID = null, string CountDes = null)
        {
            string ConnectionString2 = ConfigurationManager.ConnectionStrings["SA"].ConnectionString;
            try
            {
                string SentenciaSQL = null;
                DataSet DataSql = new DataSet();
                SqlConnection cnn = new SqlConnection(ConnectionString2);

                if (Parameter == "T")
                {
                    SentenciaSQL = string.Format("SELECT * FROM {0}", Table);
                }
                else if (Parameter == "P")
                {
                    SentenciaSQL = string.Format("SELECT * FROM {0} WHERE plans={1} AND suite={2} AND CaseID={3}", Table, plan, suite, CaseID);
                }
                else if (Parameter == "U")
                {
                    SentenciaSQL = string.Format("DELETE FROM {0} WHERE plans = {1} AND suite = {2} AND CaseID={3}", Table, plan, suite, CaseID);
                }
                else if (Parameter == "UP")
                {
                    SentenciaSQL = string.Format("UPDATE {0} SET CountDes={1} FROM {2} WHERE plans = {3} AND suite = {4} AND CaseID={5}", Table, CountDes, Table, plan, suite, CaseID);
                }

                SqlCommand cmd = new SqlCommand(SentenciaSQL, cnn);
                cmd.CommandType = CommandType.Text;
                DataTable dTable = new DataTable("dTable");
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dTable);
                DataSql.Tables.Add(dTable);
                return DataSql;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return null;
            }
        }

        public static void ExecuteSentence(string Sentence, string database, string user)
        {
            try
            {
                string ConnectionString2 = ConfigurationManager.ConnectionStrings[user].ConnectionString;
                if (database == "SQL")
                {
                    SqlConnection cnn = new SqlConnection(ConnectionString2);
                    cnn.Open();
                    SqlCommand command = cnn.CreateCommand();
                    SqlTransaction transaction;
                    transaction = cnn.BeginTransaction("Update DigiFlag y Dia 31");
                    command.Connection = cnn;
                    command.Transaction = transaction;
                    try
                    {
                        command.CommandText = Sentence;
                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        Console.WriteLine(ex.ToString());
                    }
                    cnn.Close();
                }
                else if (database == "ORA")
                {
                    OracleConnection cnn = new OracleConnection(ConnectionString2);
                    cnn.Open();
                    OracleCommand command = cnn.CreateCommand();
                    OracleTransaction transaction;
                    transaction = cnn.BeginTransaction(IsolationLevel.ReadCommitted);
                    command.Transaction = transaction;
                    try
                    {
                        command.CommandText = Sentence;
                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        Console.WriteLine(e.ToString());
                    }
                    cnn.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}

