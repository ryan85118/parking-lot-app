using Microsoft.Win32;
using System;
using System.Collections.Specialized;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace parking_lot_app.Model.MyView
{
    class File
    {
        private static readonly string eventFilePath = "./event";
        private static readonly string eventFileName = "event.txt";
        private static readonly string eventFile = Path.Combine(eventFilePath, eventFileName);
        private static readonly string outputFilePath = "./output";
        private string tempPath = null;
        private string tempName = null;

        public File()
        {
            Directory.CreateDirectory(eventFilePath);
        }

        public void OpenFile()
        {
            OpenFileDialog odXls = new OpenFileDialog();
            //指定相應的開啟文件的目錄  AppDomain.CurrentDomain.BaseDirectory定位到Debug目錄，再根據實際情況進行目錄調整
            string folderPatha = AppDomain.CurrentDomain.BaseDirectory + @"databackup\";
            odXls.InitialDirectory = folderPatha;
            odXls.Filter = "所有 Excel 檔案(*.xl*)|*.xl*|All files (*.*)|*.*";
            odXls.RestoreDirectory = true;
            if (odXls.ShowDialog() == true)
            {
                using (StreamWriter sw = new StreamWriter(eventFile, true))
                {
                    sw.WriteLine(odXls.FileName);
                    tempPath = odXls.FileName;
                    tempName = Path.GetFileNameWithoutExtension(tempPath);
                    sw.WriteLine(tempName);
                }
                ReadFile();
            }
        }

        private void ReadFile()
        {
            Directory.CreateDirectory(outputFilePath);
            OleDbConnection ole = null;
            //OleDbConnection connection = null;
            string strConn = Path.GetExtension(tempPath.Trim()).ToLower() == ".xl"
                ? $"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={tempPath.Trim()}; Extended Properties=Excel 5.0;Persist Security Info=False"
                 : $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={tempPath.Trim()}; Extended Properties='Excel 8.0; HDR=NO; IMEX=1'";

            StringCollection sheets = ExcelSheetNames();
            foreach (string s in sheets)
            {
                using (StreamWriter sw = new StreamWriter(eventFile, true))
                {
                    sw.WriteLine("sheet: " + s);
                }
            }

            string firstSheet = sheets[0];

            // connection = new OleDbConnection(strConn);


            /*  string queryString = $"SELECT LENGTH('F1') from user_tab_columns where table_name =  [{firstSheet}] ;";
              using (OleDbConnection connection = new OleDbConnection(strConn))
              {
                  OleDbCommand command = new OleDbCommand(queryString, connection);
                  connection.Open();
                  OleDbDataReader reader = command.ExecuteReader();

                  while (reader.Read())
                  {
                      Console.WriteLine(reader.GetInt32(0) + ", " + reader.GetString(1));
                  }
                  // always call Close when done reading.
                  reader.Close();
              }*/
            string strExcel = $"select * from [{firstSheet}]";


            // Console.WriteLine(column.);


            try
            {
                ole = new OleDbConnection(strConn);
                ole.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(strExcel, ole);
                ole.Close();
                DataTable dt = new DataTable();
                da.Fill(dt);
                string filePatha = Path.Combine(outputFilePath, tempName + ".csv");
                SaveToCSV(dt, filePatha);
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
                string data = string.Empty;
                StreamWriter wr = new StreamWriter(FilePatha, false, System.Text.Encoding.Default);
                foreach (DataColumn column in oTable.Columns)
                {
                    data += column.ColumnName + ",";
                }
                data += "\n";
                wr.Write(data);
                data = string.Empty;

                Console.WriteLine(oTable.Columns.Count);

                Console.WriteLine(oTable.Rows.Count);
                int rowCount = oTable.Rows.Count;
                foreach (DataColumn column in oTable.Columns)
                {
                    int dataCount = 0;
                    foreach (DataRow row in oTable.Rows)
                    {
                        if (row[column].ToString()!="")
                        {
                            dataCount++;
                        }
                    }
                    Console.WriteLine(dataCount);
                    //data += row[column].ToString().Trim() + ",";

                }



                foreach (DataRow row in oTable.Rows)
                {
                    // Console.WriteLine(row["F2"].ToString());
                    //int i = ((OleDbType)Int32.Parse(row["DATA_TYPE"].ToString()) != OleDbType.WChar) ? -1 : Int32.Parse(row["CHARACTER_MAXIMUM_LENGTH"].ToString());
                    //Console.WriteLine(i);

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

        private StringCollection ExcelSheetNames()
        {
            StringCollection names = new StringCollection();
            string strConn;
            strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + tempPath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=2'";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable sheetNames = conn.GetOleDbSchemaTable
            (System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            conn.Close();
            foreach (DataRow dr in sheetNames.Rows)
            {
                names.Add(dr[2].ToString());
            }
            return names;
        }

    }
}
