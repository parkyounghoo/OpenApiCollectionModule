using Open_Api_Collection_Module.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Open_Api_Collection_Module.Db
{
    class ModuleDb
    {
        private string connectionString = "server = localhost; uid = sa; pwd = 1111; database = PrivateData;";

        public SqlConnection GetDbConnection()
        {
            SqlConnection sqlConn = new SqlConnection(connectionString);
            sqlConn.Open();

            return sqlConn;
        }

        public void InsertApi(SqlConnection sqlConn, SqlTransaction tran, API_Model model)
        {
            SqlCommand sqlComm = new SqlCommand();
            sqlComm.Connection = sqlConn;
            sqlComm.Transaction = tran;
            sqlComm.CommandText = "insert into API_BATCH_LIST (API_NAME,API_Purpose,API_URL,API_Response,TABLE_NAME,API_Proc_Name,API_CreateDt,API_Date_YN,API_S_Date_Parameter,API_E_Date_Parameter,API_PageResult_YN,API_PageResult_Name,API_PageOfRows_Name,API_TotalCount_Name,API_Data_Collection,API_Data_Service,API_Collection_Cycle,API_Collection_Cycle_Date) " +
                "values (@API_NAME,@API_Purpose,@API_URL,@API_Response,@TABLE_NAME,@API_Proc_Name,@API_CreateDt,@API_Date_YN,@API_S_Date_Parameter,@API_E_Date_Parameter,@API_PageResult_YN,@API_PageResult_Name,@API_PageOfRows_Name,@API_TotalCount_Name,@API_Data_Collection,@API_Data_Service,@API_Collection_Cycle,@API_Collection_Cycle_Date)";

            sqlComm.Parameters.AddWithValue("@API_NAME", model.API_NAME == null ? "" : model.API_NAME);
            sqlComm.Parameters.AddWithValue("@API_Purpose", model.API_Purpose == null ? "" : model.API_Purpose);
            sqlComm.Parameters.AddWithValue("@API_URL", model.API_URL == null ? "" : model.API_URL);
            sqlComm.Parameters.AddWithValue("@API_Response", model.API_Response == null ? "" : model.API_Response);
            sqlComm.Parameters.AddWithValue("@TABLE_NAME", model.TABLE_NAME == null ? "" : model.TABLE_NAME);
            sqlComm.Parameters.AddWithValue("@API_Proc_Name", model.API_Proc_Name == null ? "" : model.API_Proc_Name);
            sqlComm.Parameters.AddWithValue("@API_CreateDt", model.API_CreateDt == null ? "" : model.API_CreateDt);
            sqlComm.Parameters.AddWithValue("@API_Date_YN", model.API_Date_YN == null ? "" : model.API_Date_YN);
            sqlComm.Parameters.AddWithValue("@API_S_Date_Parameter", model.API_S_Date_Parameter == null ? "" : model.API_S_Date_Parameter);
            sqlComm.Parameters.AddWithValue("@API_E_Date_Parameter", model.API_E_Date_Parameter == null ? "" : model.API_E_Date_Parameter);
            sqlComm.Parameters.AddWithValue("@API_PageResult_YN", model.API_PageResult_YN == null ? "" : model.API_PageResult_YN);
            sqlComm.Parameters.AddWithValue("@API_PageResult_Name", model.API_PageResult_Name == null ? "" : model.API_PageResult_Name);
            sqlComm.Parameters.AddWithValue("@API_PageOfRows_Name", model.API_PageOfRows_Name == null ? "" : model.API_PageOfRows_Name);
            sqlComm.Parameters.AddWithValue("@API_TotalCount_Name", model.API_TotalCount_Name == null ? "" : model.API_TotalCount_Name);
            sqlComm.Parameters.AddWithValue("@API_Data_Collection", model.API_Data_Collection == null ? "" : model.API_Data_Collection);
            sqlComm.Parameters.AddWithValue("@API_Data_Service", model.API_Data_Service == null ? "" : model.API_Data_Service);
            sqlComm.Parameters.AddWithValue("@API_Collection_Cycle", model.API_Collection_Cycle == null ? "" : model.API_Collection_Cycle);
            sqlComm.Parameters.AddWithValue("@API_Collection_Cycle_Date", model.API_CreateDt == null ? "" : model.API_CreateDt);
            
            sqlComm.ExecuteNonQuery();
        }

        public void CreateApiTable(SqlConnection sqlConn, SqlTransaction tran, API_Model model)
        {
            SqlCommand sqlComm = new SqlCommand();
            sqlComm.Connection = sqlConn;
            sqlComm.Transaction = tran;
            sqlComm.CommandText = model.API_CREATE_TABLE;
            sqlComm.ExecuteNonQuery();
        }

        public void CreateApiProc(SqlConnection sqlConn, SqlTransaction tran, API_Model model)
        {
            SqlCommand sqlComm = new SqlCommand();
            sqlComm.Connection = sqlConn;
            sqlComm.Transaction = tran;
            sqlComm.CommandText = model.API_CREATE_PROC;
            sqlComm.ExecuteNonQuery();
        }

        public bool getTableList(string tableName)
        {
            string queryString = "select * from API_BATCH_LIST where TABLE_NAME = '" + tableName + "'";
            using (SqlConnection sqlConn = new SqlConnection(connectionString))
            {
                DataSet ds = new DataSet();
                SqlDataAdapter _SqlDataAdapter = new SqlDataAdapter();
                _SqlDataAdapter.SelectCommand = new SqlCommand(queryString, sqlConn);
                _SqlDataAdapter.Fill(ds);

                if (ds.Tables[0].Rows.Count == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public void InsertApiDescription(SqlConnection sqlConn, SqlTransaction tran, List<TableModel> list, string tableName)
        {
            SqlCommand sqlComm = new SqlCommand();
            sqlComm.Connection = sqlConn;
            sqlComm.Transaction = tran;

            sqlComm.CommandText = "insert into API_DESCRIPTION (TABLE_NAME,ColumnName,ColumnType,ColumnSize,ColumnDescription) " +
                "values (@TABLE_NAME,@ColumnName,@ColumnType,@ColumnSize,@ColumnDescription)";

            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].ColumnName != "" || list[i].ColumnSize != "")
                {
                    sqlComm.Parameters.Clear();

                    TableModel model = list[i];
                    sqlComm.Parameters.AddWithValue("@TABLE_NAME", tableName == null ? "" : tableName);
                    sqlComm.Parameters.AddWithValue("@ColumnName", model.ColumnName == null ? "" : model.ColumnName);
                    sqlComm.Parameters.AddWithValue("@ColumnType", model.ColumnType == null ? "" : model.ColumnType);
                    sqlComm.Parameters.AddWithValue("@ColumnSize", model.ColumnSize == null ? "" : model.ColumnSize);
                    sqlComm.Parameters.AddWithValue("@ColumnDescription", model.ColumnDescription == null ? "" : model.ColumnDescription);

                    sqlComm.ExecuteNonQuery();
                }
            }
        }

        public DataSet getApiList()
        {
            string queryString = "select * from API_BATCH_LIST";
            using (SqlConnection sqlConn = new SqlConnection(connectionString))
            {
                DataSet ds = new DataSet();
                SqlDataAdapter _SqlDataAdapter = new SqlDataAdapter();
                _SqlDataAdapter.SelectCommand = new SqlCommand(queryString, sqlConn);
                _SqlDataAdapter.Fill(ds);

                return ds;
            }
        }

        public DataSet getApiDescription(string tableName)
        {
            string queryString = "select * from API_DESCRIPTION where TABLE_NAME = '" + tableName + "'";
            using (SqlConnection sqlConn = new SqlConnection(connectionString))
            {
                DataSet ds = new DataSet();
                SqlDataAdapter _SqlDataAdapter = new SqlDataAdapter();
                _SqlDataAdapter.SelectCommand = new SqlCommand(queryString, sqlConn);
                _SqlDataAdapter.Fill(ds);

                return ds;
            }
        }

        public void InsertApiList(SqlConnection sqlConn, SqlTransaction tran, Dictionary<string, List<string>> list, string tableName, List<TableModel> list2)
        {
            SqlCommand sqlComm = new SqlCommand();
            sqlComm.Connection = sqlConn;
            sqlComm.Transaction = tran;


            int cnt = 0;
            for (int i = 0; i < list2.Count; i++)
            {
                if (cnt < list[list2[i].ColumnName].Count)
                {
                    cnt = list[list2[i].ColumnName].Count;
                }
            }

            for (int i = 0; i < cnt; i++)
            {
                sqlComm.Parameters.Clear();

                StringBuilder sb = new StringBuilder();
                sb.Append("insert into " + tableName + " values (");

                for (int j = 0; j < list2.Count; j++)
                {
                    if (j == 0)
                    {
                        if (list[list2[j].ColumnName].Count == 0)
                        {
                            sb.Append("''");
                        }
                        else
                        {
                            string value = list[list2[j].ColumnName].ElementAt(i) == null ? "" : list[list2[j].ColumnName].ElementAt(i);
                            sb.Append("'" + value + "'");
                        }
                    }
                    else
                    {
                        if (list[list2[j].ColumnName].Count == 0)
                        {
                            sb.Append(",''");
                        }
                        else
                        {
                            string value = list[list2[j].ColumnName].ElementAt(i) == null ? "" : list[list2[j].ColumnName].ElementAt(i);
                            sb.Append(",'" + value + "'");
                        }
                    }
                }

                sb.Append(")");

                sqlComm.CommandText = sb.ToString();

                sqlComm.ExecuteNonQuery();
            }
        }
    }
}
