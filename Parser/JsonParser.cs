using Newtonsoft.Json.Linq;
using Open_Api_Collection_Module.Model;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Open_Api_Collection_Module.Parser
{
    internal class JsonParser
    {
        public string getJson(string url)
        {
            using (WebClient wc = new WebClient())
            {
                wc.Encoding = Encoding.UTF8;
                string json = wc.DownloadString(url);
                return json;
            }
        }

        public Dictionary<string, List<string>> getJsonList(string json, List<TableModel> fieldList)
        {
            JObject outer = JObject.Parse(json);
            List<List<TableModel>> list = new List<List<TableModel>>();

            List<TableModel> TableList = new List<TableModel>();
            foreach (var item in outer.Descendants())
            {
                for (int i = 0; i < fieldList.Count; i++)
                {
                    if (item.Path.Substring(item.Path.LastIndexOf('.') + 1) == fieldList[i].ColumnName)
                    {
                        if (item.Type.ToString() == "Property")
                        {
                            TableModel model = new TableModel();
                            model.ColumnName = fieldList[i].ColumnName;
                            model.Value = item.Value<JProperty>().Value.ToString();

                            TableList.Add(model);
                        }
                    }
                }
            }

            Dictionary<string, List<string>> dicJson = new Dictionary<string, List<string>>();

            for (int i = 0; i < fieldList.Count; i++)
            {
                IEnumerable<string> itemList = from model in TableList
                                               where model.ColumnName == fieldList[i].ColumnName
                                               select model.Value;

                dicJson.Add(fieldList[i].ColumnName, itemList.ToList());
            }

            return dicJson;

            #region 주석
            //for (int k = 0; k < dicJson[fieldList[0].ColumnName].Count; k++)
            //{
            //    for (int j = 0; j < fieldList.Count; j++)
            //    {
            //        string test = dicJson[fieldList[j].ColumnName].ElementAt(k);
            //        //if (j == 0)
            //        //{
            //        //    sb.Append("'" + list[list2[j].ColumnName].ToString() + "'");
            //        //}
            //        //else
            //        //{
            //        //    sb.Append(",'" + list[i][j].Value.ToString() + "'");
            //        //}
            //    }
            //}
            //int cnt = 0;
            //for (int i = 0; i < TableList.Count / fieldList.Count; i++)
            //{
            //    List<TableModel> modelList = new List<TableModel>();

            //    for (int j = 0; j < fieldList.Count; j++)
            //    {
            //        TableModel model = new TableModel();
            //        model.ColumnName = fieldList[j].ColumnName;
            //        model.Value = TableList[cnt].Value;
            //        modelList.Add(model);
            //        cnt++;
            //    }

            //    list.Add(modelList);
            //}



            //for (int i = 0; i < fieldList.Count; i++)
            //{
            //    List<TableModel> TableList = new List<TableModel>();

            //    foreach (var item in outer.Descendants())
            //    {
            //        if (item.Path.Contains(fieldList[i].ColumnName))
            //        {
            //            if (item.Type.ToString() == "Property")
            //            {
            //                TableModel model = new TableModel();
            //                model.ColumnName = fieldList[i].ColumnName;
            //                model.Value = item.Value<JProperty>().Value.ToString();

            //                TableList.Add(model);
            //            }
            //        }
            //    }

            //    list.Add(TableList);
            //}

            //JToken jToken = JToken.Parse(json);
            //JObject inner = outer["bidNtceNo"].Value<JObject>();
            //JArray array = JArray.Parse(json);
            //JObject array = JObject.Parse(json);
            //string keys = array.Properties().Select(p => "bidNtceNo").First();
            //List<List<string>> list = new List<List<string>>();

            //foreach (JObject jobj in array)
            //{
            //    List<string> model = new List<string>();
            //    for (int i = 0; i < fieldList.Count; i++)
            //    {
            //        model.Add(jobj[fieldList[i]].ToString());
            //    }

            //    list.Add(model);
            //}

            //foreach (JToken item in array["data"])
            //{
            //    List<string> model = new List<string>();
            //    for (int i = 0; i < fieldList.Count; i++)
            //    {
            //        model.Add(item[fieldList[i]].ToString());
            //    }

            //    list.Add(model);
            //}

            #endregion 주석
        }
    }
}