using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace userDeal
{
    public class CsvHelper
    {
        public bool OutToCSV(string title, System.Data.DataTable dtinfo)
        {
            int titleColumnSpan = dtinfo.Columns.Count;//标题的跨列数
            string fileName = "";//保存的excel文件名
            if (dtinfo.Rows.Count == 0)
                return false;
            /*保存对话框*/
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "导出CSV文件|*.CSV";
            sfd.FileName = title + DateTime.Now.ToString("yyyyMMddhhmmss");

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                fileName = sfd.FileName;
                SaveCSV(dtinfo, fileName);
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 将DataTable中数据写入到CSV文件中
        /// </summary>
        /// <param name="dt">提供保存数据的DataTable</param>
        /// <param name="fileName">CSV的文件路径</param>
        public void SaveCSV(DataTable dt, string fileName)
        {
            FileStream fs = new FileStream(fileName, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";

            //写出列名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                data += dt.Columns[i].ColumnName.ToString();
                if (i < dt.Columns.Count - 1)
                {
                    data += ",";
                }
            }
            sw.WriteLine(data);

            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    data += "\"\t" + (dt.Rows[i][j]).ToString().Replace("\"", "\"\"") + "\"";
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }

            sw.Close();
            fs.Close();
            MessageBox.Show("CSV文件保存成功！");
        }
    }
}
