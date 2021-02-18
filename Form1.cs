using Microsoft.SqlServer.TransactSql.ScriptDom;
using Open_Api_Collection_Module.Db;
using Open_Api_Collection_Module.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Open_Api_Collection_Module
{
    public partial class Form1 : Form
    {
        private const int WM_COPYDATA = 0x4A;

        public struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;

            [MarshalAs(UnmanagedType.LPStr)]
            public string lpData;
        }

        public Form1()
        {
            InitializeComponent();
            responseButtonJ.Checked = true;
            rbDay.Checked = true;

            #region 테스트용 주석

            //////배치 API 테이블 호출
            //ModuleDb db = new ModuleDb();
            //DataSet ds = db.getApiList();

            //List<API_Model> list = new List<API_Model>();
            //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            //{
            //    API_Model model = new API_Model();
            //    model.API_NAME = ds.Tables[0].Rows[i]["API_NAME"].ToString();
            //    model.API_Purpose = ds.Tables[0].Rows[i]["API_Purpose"].ToString();
            //    model.API_URL = ds.Tables[0].Rows[i]["API_URL"].ToString();
            //    model.API_Response = ds.Tables[0].Rows[i]["API_Response"].ToString();
            //    model.API_Table_Name = ds.Tables[0].Rows[i]["API_Table_Name"].ToString();
            //    model.API_Proc_Name = ds.Tables[0].Rows[i]["API_Proc_Name"].ToString();
            //    model.API_CreateDt = ds.Tables[0].Rows[i]["API_CreateDt"].ToString();
            //    model.API_Date_YN = ds.Tables[0].Rows[i]["API_Date_YN"].ToString();
            //    model.API_S_Date_Parameter = ds.Tables[0].Rows[i]["API_S_Date_Parameter"].ToString();
            //    model.API_E_Date_Parameter = ds.Tables[0].Rows[i]["API_E_Date_Parameter"].ToString();
            //    model.API_PageResult_YN = ds.Tables[0].Rows[i]["API_PageResult_YN"].ToString();
            //    model.API_PageResult_Name = ds.Tables[0].Rows[i]["API_PageResult_Name"].ToString();
            //    model.API_PageOfRows_Name = ds.Tables[0].Rows[i]["API_PageOfRows_Name"].ToString();
            //    model.API_TotalCount_Name = ds.Tables[0].Rows[i]["API_TotalCount_Name"].ToString();
            //    model.API_Data_Collection = ds.Tables[0].Rows[i]["API_Data_Collection"].ToString();
            //    model.API_Data_Service = ds.Tables[0].Rows[i]["API_Data_Service"].ToString();
            //    model.API_Collection_Cycle = ds.Tables[0].Rows[i]["API_Collection_Cycle"].ToString();
            //    model.API_Collection_Cycle_Date = ds.Tables[0].Rows[i]["API_Collection_Cycle_Date"].ToString();

            //    list.Add(model);
            //}

            //CollectionModule module = new CollectionModule();

            //for (int i = 0; i < list.Count; i++)
            //{
            //    if (list[i].API_Table_Name == "TEST2")
            //    {
            //        List<TableModel> listTable = new List<TableModel>();
            //        DataSet dsDescription = db.getApiDescription(list[i].API_Table_Name);
            //        for (int j = 0; j < dsDescription.Tables[0].Rows.Count; j++)
            //        {
            //            DataRow dr = dsDescription.Tables[0].Rows[j];
            //            TableModel model = new TableModel();
            //            model.ColumnName = dr["ColumnName"].ToString();

            //            listTable.Add(model);
            //        }

            //        module.apiRequest(list[i], listTable);
            //    }
            //}

            #endregion 테스트용 주석
        }

        protected override void WndProc(ref Message m)
        {
            try
            {
                switch (m.Msg)
                {
                    case WM_COPYDATA:
                        COPYDATASTRUCT cds = (COPYDATASTRUCT)m.GetLParam(typeof(COPYDATASTRUCT));
                        SetTableTemplate(cds.lpData);
                        break;

                    default:
                        base.WndProc(ref m);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void SetTableTemplate(string result)
        {
            string[] split = result.Split('%');

            List<TableModel> list = new List<TableModel>();
            for (int i = 0; i < (split.Length - 1) / 3; i++)
            {
                if (split[(i * 3) + 1].ToString() != "")
                {
                    TableModel model = new TableModel();
                    model.ColumnName = split[(i * 3) + 1].ToString();
                    model.ColumnSize = split[(i * 3) + 2].ToString();
                    model.ColumnDescription = split[(i * 3) + 3].ToString();

                    list.Add(model);
                }
            }

            tlpTable.Visible = true;
            lbTable.Visible = true;
            btnAddRow.Visible = true;
            if (list.Count >= 13)
            {
                for (int i = 0; i < list.Count - 13; i++)
                {
                    Row_Add();
                }
            }

            for (int i = 0; i < list.Count; i++)
            {
                TableModel model = list[i];

                tlpTable.Controls.Find("tbColumnName" + ((i * 4) + 3).ToString(), true)[0].Text = model.ColumnName;
                tlpTable.Controls.Find("tbColumnName" + ((i * 4) + 5).ToString(), true)[0].Text = model.ColumnSize;
                tlpTable.Controls.Find("tbColumnName" + ((i * 4) + 6).ToString(), true)[0].Text = model.ColumnDescription;
            }
        }

        private void btnbatch_Click(object sender, EventArgs e)
        {
            bool check = true;

            rtbResult.Text = "";

            #region 유효성 체크

            //유효성 체크
            if (tbOpenApiName.Text.Trim() == "")
            {
                rtbResult.Text += "오픈 API 명을 입력해 주세요.\n";

                check = false;
            }
            if (tbPurpose.Text.Trim() == "")
            {
                rtbResult.Text += "사용목적을 입력해 주세요.\n";

                check = false;
            }
            if (tbUrl.Text.Trim() == "")
            {
                rtbResult.Text += "호출 URL을 입력해 주세요.\n";

                check = false;
            }
            else if (!CheckURLValid(tbUrl.Text))
            {
                rtbResult.Text += "URL 형식이 아닙니다.\n";

                check = false;
            }

            if (tbTableName.Text == "")
            {
                rtbResult.Text += "테이블명을 입력해 주세요.\n";

                check = false;
            }

            if (ckbDate.Checked)
            {
                if (tbSDateParameter.Text.Trim() == "")
                {
                    rtbResult.Text += "날짜 Parameter를 입력해 주세요.\n";

                    check = false;
                }
                if (tbEDateParameter.Text.Trim() == "")
                {
                    rtbResult.Text += "날짜 Format을 입력해 주세요.\n";

                    check = false;
                }
            }

            if (ckbPageResult.Checked)
            {
                if (tbPageResultName.Text.Trim() == "")
                {
                    rtbResult.Text += "페이지 결과 수 Parameter를 입력해 주세요.\n";

                    check = false;
                }
                if (tbTotalCountName.Text.Trim() == "")
                {
                    rtbResult.Text += "전체 결과수 Parameter를 입력해 주세요.\n";

                    check = false;
                }
            }

            #endregion 유효성 체크

            if (check)
            {
                CollectionModule module = new CollectionModule();

                API_Model model = new API_Model();
                model.API_NAME = tbOpenApiName.Text;
                model.API_Purpose = tbPurpose.Text;
                model.API_URL = tbUrl.Text;
                model.API_Response = responseButtonJ.Checked ? "json" : "xml";
                model.TABLE_NAME = tbTableName.Text;
                model.API_Proc_Name = "proc_" + tbTableName.Text;
                model.API_CreateDt = DateTime.Now.ToString("yyyy-MM-dd");
                model.API_Data_Collection = tbDataCollection.Text;
                model.API_Data_Service = tbDataService.Text;

                if (cbBatchYn.Checked)
                {
                    model.API_PageResult_YN = ckbPageResult.Checked ? "Y" : "N";
                    model.API_Date_YN = ckbDate.Checked ? "Y" : "N";
                    model.API_S_Date_Parameter = tbSDateParameter.Text;
                    model.API_E_Date_Parameter = tbEDateParameter.Text;
                    model.API_PageResult_Name = tbPageResultName.Text;
                    model.API_PageOfRows_Name = tbPageOfRows.Text;
                    model.API_TotalCount_Name = tbTotalCountName.Text;
                    model.API_Collection_Cycle = rbDay.Checked ? "D" : rbMonth.Checked ? "M" : rbYear.Checked ? "Y" : "";
                }
                else
                {
                    model.API_PageResult_YN = "";
                    model.API_Date_YN = "";
                    model.API_S_Date_Parameter = "";
                    model.API_E_Date_Parameter = "";
                    model.API_PageResult_Name = "";
                    model.API_PageOfRows_Name = "";
                    model.API_TotalCount_Name = "";
                    model.API_Collection_Cycle = "";
                }

                List<TableModel> list = TablePanelToTableModel();
                string sqlmessage = "";

                //테이블 생성
                string tableString = module.PanelCreateTable(list, model.TABLE_NAME);
                if (!ValidateSql(tableString, out sqlmessage))
                {
                    rtbResult.Text = sqlmessage;

                    return;
                }

                //프로시저 생성
                string procString = module.PanelCreateProc(list, model.API_Proc_Name, model.TABLE_NAME);
                if (!ValidateSql(tableString, out sqlmessage))
                {
                    rtbResult.Text = sqlmessage;

                    return;
                }

                model.API_CREATE_TABLE = tableString;
                model.API_CREATE_PROC = procString;

                string message = "";
                module.setCollectionModule(model, list, out message);

                rtbResult.Text = message;
            }
        }

        /// <summary>
        /// url 유효성 검사
        /// </summary>
        /// <param name="source">url 값</param>
        /// <returns></returns>
        private bool CheckURLValid(string source)
        {
            Uri uriResult;
            return Uri.TryCreate(source, UriKind.Absolute, out uriResult);
        }

        /// <summary>
        /// 쿼리 유효성 검사
        /// </summary>
        /// <param name="str">쿼리 값</param>
        /// <returns></returns>
        private bool ValidateSql(string str, out string message)
        {
            if (string.IsNullOrWhiteSpace(str))
            {
                message = "테이블명을 입력하세요.";

                return false;
            }
            var parser = new TSql120Parser(false);
            IList<ParseError> errors;
            using (var reader = new StringReader(str))
            {
                parser.Parse(reader, out errors);
            }

            if (errors.Count == 0)
            {
                message = "";

                return true;
            }
            else
            {
                message = "";
                for (int i = 0; i < errors.Count; i++)
                {
                    message += errors[i].Message + "/";
                }

                return false;
            }
        }

        private void btnTableName_Click(object sender, EventArgs e)
        {
            if (tbTableName.Text.Trim() == "")
            {
                rtbResult.Text = "테이블 명을 입력해 주세요.\n";
            }
            else
            {
                CollectionModule module = new CollectionModule();
                if (module.getTableNameCheck(tbTableName.Text.Trim()))
                {
                    tlpTable.Visible = true;
                    lbTable.Visible = true;
                    btnAddRow.Visible = true;
                    if (tlpTable.RowStyles.Count == 1)
                    {
                        tlpTable.Visible = true;
                        lbTable.Visible = true;
                        btnAddRow.Visible = true;
                    }
                }
                else
                {
                    rtbResult.Text = "같은 테이블명이 존재 합니다.\n";
                    tbTableName.Text = "";
                }
            }
        }

        private void ckbDate_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                tbSDateParameter.Enabled = true;
                tbEDateParameter.Enabled = true;
            }
            else
            {
                tbSDateParameter.Enabled = false;
                tbEDateParameter.Enabled = false;
            }
        }

        private void ckbPageResult_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                tbPageResultName.Enabled = true;
                tbPageOfRows.Enabled = true;
                tbTotalCountName.Enabled = true;
            }
            else
            {
                tbPageResultName.Enabled = false;
                tbPageOfRows.Enabled = false;
                tbTotalCountName.Enabled = false;
            }
        }

        private void Row_Add()
        {
            tlpTable.RowCount = tlpTable.RowCount + 1;
            tlpTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tlpTable.Controls.Add(new TextBox() { Name = "tbColumnName" + (tlpTable.Controls.Count - 1).ToString(), Anchor = AnchorStyles.None, Width = 109 }, 0, tlpTable.RowCount - 1);
            tlpTable.Controls.Add(new TextBox() { Name = "tbColumnName" + (tlpTable.Controls.Count - 1).ToString(), Anchor = AnchorStyles.None, Text = "VARCHAR", Enabled = false }, 1, tlpTable.RowCount - 1);
            tlpTable.Controls.Add(new TextBox() { Name = "tbColumnName" + (tlpTable.Controls.Count - 1).ToString(), Anchor = AnchorStyles.None }, 2, tlpTable.RowCount - 1);
            tlpTable.Controls.Add(new TextBox() { Name = "tbColumnName" + (tlpTable.Controls.Count - 1).ToString(), Anchor = AnchorStyles.None, Width = 238 }, 3, tlpTable.RowCount - 1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tlpTable.Visible)
            {
                Row_Add();
            }
        }

        private List<TableModel> TablePanelToTableModel()
        {
            List<TableModel> tableList = new List<TableModel>();

            List<string> list = new List<string>();
            for (int i = 0; i < tlpTable.Controls.Count - 4; i++)
            {
                list.Add(tlpTable.Controls.Find("tbColumnName" + (i + 3).ToString(), true)[0].Text);
            }

            for (int j = 0; j < list.Count / 4; j++)
            {
                if (list[(j * 4)].ToString() != "")
                {
                    TableModel model = new TableModel();
                    model.ColumnName = list[(j * 4)].ToString();
                    model.ColumnType = list[(j * 4) + 1].ToString();
                    model.ColumnSize = list[(j * 4) + 2].ToString();
                    model.ColumnDescription = list[(j * 4) + 3].ToString();

                    tableList.Add(model);
                }
            }

            return tableList;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern System.IntPtr FindWindow(string lpClassName, string lpWindowName);

        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("엑셀을 실행 할 수 없습니다.");
            }

            // 엑셀 시트 만들기
            Microsoft.Office.Interop.Excel.Workbooks workbooks = excelApp.Workbooks;
            Microsoft.Office.Interop.Excel._Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Sheets sheets = workbook.Worksheets;

            // 엑셀 보이기
            string caption = excelApp.Caption;
            IntPtr handler = FindWindow(null, caption);
            SetForegroundWindow(handler);
            excelApp.Visible = true;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Form1 NewForm = new Form1();
            NewForm.Show();
            this.Dispose(false);
        }

        private void cbBatchYn_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked)
            {
                plBatch.Enabled = true;
            }
            else
            {
                plBatch.Enabled = false;
            }
        }
    }
}