using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace userDeal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnDeal_Click(object sender, EventArgs e)
        {
             string path = "";
             OpenFileDialog openFileDialog = new OpenFileDialog();
             openFileDialog.Title = "选择文件";//选择框名称        
             openFileDialog.Filter = "xls files (*.xls)|*.xls";//选择文件的类型为Xls表格  
             if (openFileDialog.ShowDialog() == DialogResult.OK)//当点击确定               
             {
                 path = openFileDialog.FileName.Trim();//文件路径
                 //path = path.Replace("\\", "/");
                 
                 try
                 {
                     DataTable dt = getDataTable(path);
                     outCsv(dt);      
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show(ex.Message);
                 }
             }


        }

        private void outCsv(DataTable dt)
        {
            int yearNow=DateTime.Now.Year;
            String ylzhName = dt.Columns[0].ColumnName;
            String phoneName = dt.Columns[14].ColumnName;
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("数据不能为空");
                return;
            }
            DataTable dtOut = new DataTable();
            dtOut.Columns.Add("户主姓名");
            dtOut.Columns.Add("性别");
            dtOut.Columns.Add("年龄");
            dtOut.Columns.Add("身份证号");
            dtOut.Columns.Add("联系电话");
            dtOut.Columns.Add("几组");
            dtOut.Columns.Add("亲属姓名");
            dtOut.Columns.Add("亲属关系");
            dtOut.Columns.Add("亲属身份证号");
            dtOut.Columns.Add("职业");
            dtOut.Columns.Add("联系方式");
            foreach(DataRow dr in dt.Rows){
                if (dr[0] == null || dr[0].ToString() == "" ||  !dr[0].ToString().StartsWith("4"))
                {
                    continue;
                }
                DataRow drNew = dtOut.NewRow();
                if (dr[1] != null && dr[1].ToString() != "")
                {
                    drNew["几组"] = dr[1].ToString();
                }                
                if (dr[2] != null && dr[2].ToString() != "")
                {
                    drNew["户主姓名"] = dr[2].ToString();
                    if (dr[4] != null && dr[4].ToString() != "")
                    {
                        drNew["性别"] = dr[4].ToString();
                    }
                    if (dr[5] != null && dr[5].ToString() != "")
                    {
                        drNew["身份证号"] = dr[5].ToString();
                        if(dr[5].ToString().Length>=18){
                            String yearStr = dr[5].ToString().Substring(6, 4);
                            int year;
                            if (int.TryParse(yearStr, out year))
                            {
                                drNew["年龄"] = yearNow - year;
                            }
                        }
                    }
                    if (dr[14] != null && dr[14].ToString() != "")
                    {
                        drNew["联系电话"] = dr[14].ToString();
                    }
                    else
                    {
                        if (dr[0] != null && dr[0].ToString() != "")
                        {
                            String ylzh = dr[0].ToString();
                            DataRow[] rows = dt.Select(ylzhName + "= '" + ylzh + "' and " + phoneName + " is not null and " + phoneName + "  <>'' and " + phoneName + "  <>'0' ");
                            if (rows != null && rows.Length > 0)
                            {
                                drNew["联系电话"] = rows[0][14].ToString();
                            }
                        }    
                    }
                }
                else
                {
                    if (dr[3] != null && dr[3].ToString() != "")
                    {
                        drNew["亲属姓名"] = dr[3].ToString();
                    }
                    if (dr[5] != null && dr[5].ToString() != "")
                    {
                        drNew["亲属身份证号"] = dr[5].ToString();
                    }
                    if (dr[6] != null && dr[6].ToString() != "")
                    {
                        drNew["亲属关系"] = dr[6].ToString();
                    }
                    if (dr[14] != null && dr[14].ToString() != "")
                    {
                        drNew["联系方式"] = dr[14].ToString();
                    }
                }
                dtOut.Rows.Add(drNew);
            }
            new CsvHelper().OutToCSV("居民信息",dtOut);
        }       
            DataTable getDataTable(String path)
        {
            string strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\";data source=" + path + ";";
            OleDbConnection myConn = new OleDbConnection(strCon);
            myConn.Open();
            DataTable dtSheetName = myConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            String tableName = dtSheetName.Rows[0]["TABLE_NAME"].ToString();
            string strCom = " select * from [" + tableName + "]";
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            DataTable dt = new DataTable();
            myCommand.Fill(dt);
            return dt;
        }
    }
}
