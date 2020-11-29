using Microsoft.Win32;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace parking_lot_app.Model.MyView
{
    class File
    {
        string Patha = null;
        static readonly string eventFilePath = "./event/";
        static readonly string eventFileName = "event.txt";
        string eventFile = Path.Combine(eventFilePath, eventFileName);

        string tempPath = null;
        string tempName = null;

        public File(string p)
        {
            Patha = p;
        }

        public void OpenFile()
        {
            Directory.CreateDirectory(eventFilePath);
            OpenFileDialog odXls = new OpenFileDialog();
            //指定相應的開啟文件的目錄  AppDomain.CurrentDomain.BaseDirectory定位到Debug目錄，再根據實際情況進行目錄調整
            string folderPatha = AppDomain.CurrentDomain.BaseDirectory + @"databackup\";
            odXls.InitialDirectory = folderPatha;
            //// 設定檔案格式  
            //odXls.Filter = "Excel files office2003(*.xls)|*.xls|Excel office2010(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            //openFileDialog1.Filter = "圖片檔案(*.jpg)|*.jpg|(*.JPEG)|*.jpeg|(*.PNG)|*.png";
            odXls.Filter = "所有 Excel 檔案(*.xl*)|*.xl*|All files (*.*)|*.*";
            odXls.RestoreDirectory = true;
            //odXls.InitialDirectory = Environment.GetFolderPatha(Environment.SpecialFolder.MyDocuments);
            if (odXls.ShowDialog() == true)
            {
                using (StreamWriter sw = new StreamWriter(eventFile, true))
                {
                    sw.WriteLine(odXls.FileName);
                    tempPath = odXls.FileName;
                    tempName = Path.GetFileNameWithoutExtension(tempPath);
                    sw.WriteLine(tempName);
                }
                Button_Click();
            }
        }

        private void Button_Click()
        {
            OleDbConnection ole = null;
            OleDbDataAdapter da = null;
            DataTable dt = null;
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source={0};" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';", tempPath.Trim());
            if (Path.GetExtension(tempPath.Trim()).ToLower() == ".xl")
            {
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + tempPath.Trim() + ";Extended Properties=Excel 5.0;Persist Security Info=False";
            }
            string strExcel = "select * from [ReportDay$]";
            try
            {
                ole = new OleDbConnection(strConn);
                ole.Open();
                da = new OleDbDataAdapter(strExcel, ole);
                dt = new DataTable();
                da.Fill(dt);
                string filePatha = "./csv/" + tempName + ".csv";
                Directory.CreateDirectory("./csv/");
                SaveToCSV(dt, filePatha);
                ole.Close();
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(eventFile, true))
                {
                    sw.WriteLine("error: " + DateTime.Now.ToString() + ex.Message);
                }
            }
            finally
            {
                if (ole != null)
                    ole.Close();
            }

            void SaveToCSV(DataTable oTable, string FilePatha)
            {
                string data = "";
                StreamWriter wr = new StreamWriter(FilePatha, false, System.Text.Encoding.Default);
                foreach (DataColumn column in oTable.Columns)
                {
                    data += column.ColumnName + ",";
                }
                data += "\n";
                wr.Write(data);
                data = "";
                foreach (DataRow row in oTable.Rows)
                {
                    foreach (DataColumn column in oTable.Columns)
                    {
                        data += row[column].ToString().Trim() + ",";
                    }
                    data += "\n";
                    wr.Write(data);
                    data = "";
                }
                data += "\n";
                wr.Dispose();
                wr.Close();
            }
        }
    }
}
