using Open_Api_Collection_Module.Db;
using Open_Api_Collection_Module.Model;
using Open_Api_Collection_Module.Parser;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Open_Api_Collection_Module
{
    class CollectionModule
    {
        public void setCollectionModule(API_Model model, List<TableModel> list, out string errorMessage)
        {
            if (getTableNameCheck(model.TABLE_NAME))
            {
                ModuleDb db = new ModuleDb();

                //con 객체 얻기
                SqlConnection sqlConn = db.GetDbConnection();
                //tran 객체 생성
                SqlTransaction tran = sqlConn.BeginTransaction();
                try
                {
                    //API Insert
                    db.InsertApi(sqlConn, tran, model);
                    //테이블 생성
                    db.CreateApiTable(sqlConn, tran, model);
                    //프로시저 생성
                    db.CreateApiProc(sqlConn, tran, model);
                    //테이블 항목설명 저장
                    db.InsertApiDescription(sqlConn, tran, list, model.TABLE_NAME);
                    //최초 api 조회, 저장
                    db.InsertApiList(sqlConn, tran, apiRequest(model, list), model.TABLE_NAME, list);

                    tran.Commit();

                    errorMessage = "성공";
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    errorMessage = ex.Message;
                }
                finally
                {
                    sqlConn.Dispose();
                }
            }
            else
            {
                errorMessage = "같은 테이블명이 존재 합니다.";
            }
        }
        
        public bool getTableNameCheck(string tableName)
        {
            ModuleDb db = new ModuleDb();
            return db.getTableList(tableName);
        }

        public Dictionary<string, List<string>> apiRequest(API_Model model, List<TableModel> tableModelList)
        {
            //주소 생성
            string url = model.API_URL;
            //xml, json
            string result = "";
            //Api List
            Dictionary<string, List<string>> list = new Dictionary<string, List<string>>();
            if (model.API_Response == "json")
            {
                JsonParser json = new JsonParser();
                result = json.getJson(url);

                list = json.getJsonList(result, tableModelList);
            }
            else
            {
                XmlParser xml = new XmlParser();
                result = xml.getXml(url);

                list = xml.getXmlList(result, tableModelList);
            }

            return list;
        }

        public string PanelCreateTable(List<TableModel> list, string tableName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("CREATE TABLE " + tableName + "\n");
            sb.Append("(\n");
            for (int i = 0; i < list.Count; i++)
            {
                TableModel model = list[i];
                if (model.ColumnName != "")
                {
                    if (i == 0)
                    {
                        sb.Append(model.ColumnName + " " + model.ColumnType + "(" + model.ColumnSize + ")\n");
                    }
                    else
                    {
                        sb.Append("," + model.ColumnName + " " + model.ColumnType + "(" + model.ColumnSize + ")\n");
                    }
                }
            }
            sb.Append(")");

            return sb.ToString();
        }

        public string PanelCreateProc(List<TableModel> list, string procName, string tableName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("CREATE PROC " + procName + "\n");
            sb.Append("(\n");
            for (int i = 0; i < list.Count; i++)
            {
                TableModel model = list[i];
                if (model.ColumnName != "")
                {
                    if (i == 0)
                    {
                        sb.Append("@" + model.ColumnName + " " + model.ColumnType + "(" + model.ColumnSize + ")\n");
                    }
                    else
                    {
                        sb.Append(",@" + model.ColumnName + " " + model.ColumnType + "(" + model.ColumnSize + ")\n");
                    }
                }
            }
            sb.Append(")\n");
            sb.Append("AS\n");
            sb.Append("BEGIN\n");
            sb.Append("DECLARE @CNT INT\n");
            sb.Append("SELECT\n");
            sb.Append("@CNT = COUNT(*)\n");
            sb.Append("FROM " + tableName + "\n");
            sb.Append("WHERE \n");
            for (int j = 0; j < list.Count; j++)
            {
                TableModel model = list[j];
                if (model.ColumnName != "")
                {
                    if (j == 0)
                    {
                        sb.Append(model.ColumnName + " = @" + model.ColumnName + "\n");
                    }
                    else
                    {
                        sb.Append("AND " + model.ColumnName + " = @" + model.ColumnName + "\n");
                    }
                }
            }
            sb.Append("IF(@CNT = 0)\n");
            sb.Append("BEGIN\n");
            sb.Append("INSERT INTO " + tableName + "\n");
            sb.Append("VALUES\n");
            sb.Append("(\n");
            for (int k = 0; k < list.Count; k++)
            {
                TableModel model = list[k];
                if (model.ColumnName != "")
                {
                    if (k == 0)
                    {
                        sb.Append("@" + model.ColumnName + "\n");
                    }
                    else
                    {
                        sb.Append(",@" + model.ColumnName + "\n");
                    }
                }
            }
            sb.Append(")\n");
            sb.Append("END\n");
            sb.Append("END\n");

            return sb.ToString();
        }
    }
}
