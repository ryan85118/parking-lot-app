using Microsoft.Win32;
using System;
using System.Collections.Specialized;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace parking_lot_app.Model.MyView
{
    class File
    {
        private static readonly string eventFilePath = "./event";
        private static readonly string eventFileName = "event.txt";
        private static readonly string eventFile = Path.Combine(eventFilePath, eventFileName);
        private static readonly string outputFilePath = "./output";
        private static readonly string entryTimeFilePath = "./output/單月入場";
        private static readonly string stayTimeFilePath = "./output/單月停車";
        private static readonly string totalAmountFilePath = "./output/單月金額";
        private string entryTimeFile;
        private string stayTimeFile;
        private string totalAmountFile;

        private string tempPath = null;
        private string tempName = null;

        public File()
        {
            Directory.CreateDirectory(eventFilePath);
            Directory.CreateDirectory(entryTimeFilePath);
            Directory.CreateDirectory(stayTimeFilePath);
            Directory.CreateDirectory(totalAmountFilePath);
        }

        public void OpenFile()
        {
            OpenFileDialog odXls = new OpenFileDialog();
            //指定相應的開啟文件的目錄 AppDomain.CurrentDomain.BaseDirectory定位到Debug目錄，再根據實際情況進行目錄調整
            string folderPatha = AppDomain.CurrentDomain.BaseDirectory + @"databackup\";
            odXls.InitialDirectory = folderPatha;
            odXls.Filter = "所有 Excel 檔案(*.xl*)|*.xl*|All files (*.*)|*.*";
            odXls.RestoreDirectory = true;
            odXls.Multiselect = true;
            if (odXls.ShowDialog() == true)
            {
                foreach (String fileName in odXls.FileNames)
                {
                    using (StreamWriter sw = new StreamWriter(eventFile, true))
                    {
                        sw.WriteLine(fileName);
                        tempPath = fileName;
                        tempName = Path.GetFileNameWithoutExtension(tempPath);
                        sw.WriteLine(tempName);
                    }
                    ReadFile();
                }
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
            string strExcel = $"select * from [{firstSheet}]";

            try
            {
                ole = new OleDbConnection(strConn);
                ole.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(strExcel, ole);
                ole.Close();
                DataTable dt = new DataTable();
                da.Fill(dt);
                string tempFilePath = Path.Combine(outputFilePath, tempName + ".csv");
                entryTimeFile = Path.Combine(entryTimeFilePath, tempName + "-單月入場表.csv");
                stayTimeFile = Path.Combine(stayTimeFilePath, tempName + "-單月停車表.csv");
                totalAmountFile = Path.Combine(totalAmountFilePath, tempName + "-單月金額表.csv");
                SaveToCSV(dt, tempFilePath);
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(eventFile, true))
                {
                    sw.WriteLine("Error: Open Excel File failure, " + DateTime.Now.ToString() + ", " + ex.Message);
                }
            }
            finally
            {
                if (ole != null)
                {
                    ole.Close();
                }
            }
        }

        private void SaveToCSV(DataTable oTable, string csvFilePath)
        {
            int rowCount = oTable.Rows.Count;
            DataTable tempTable = new DataTable("Temp");
            DataTable EntryTimeTable = new DataTable("entryTime");
            DataTable StayTimeTable = new DataTable("stayTime");
            DataTable TotalAmountTable = new DataTable("totalAmount");


            //Now 停車票號,入場時間,出場時間,發票號碼,收費金額
            tempTable.Columns.Add("停車票號", typeof(string));
            tempTable.Columns.Add("入場時間", typeof(string));
            tempTable.Columns.Add("出場時間", typeof(string));
            tempTable.Columns.Add("發票號碼", typeof(string));
            tempTable.Columns.Add("收費金額", typeof(string));

            //EntryTimeTable
            EntryTimeTable.Columns.Add("日期／時間", typeof(string));
            for (int i = 0; i < 24; i++)
            {
                EntryTimeTable.Columns.Add(i.ToString(), typeof(int));
            }
            EntryTimeTable.Columns.Add("總計", typeof(int));

            //StayTimeTable
            //日期／小時 <0.5 <1小時 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 >=20 總計
            StayTimeTable.Columns.Add("日期／小時", typeof(string));
            StayTimeTable.Columns.Add("<0.5", typeof(int));
            StayTimeTable.Columns.Add("<1", typeof(int));
            for (int i = 1; i < 20; i++)
            {
                StayTimeTable.Columns.Add(i.ToString(), typeof(int));
            }
            StayTimeTable.Columns.Add(">=20", typeof(int));
            StayTimeTable.Columns.Add("總計", typeof(int));

            //TotalAmountTable
            TotalAmountTable.Columns.Add("日期／元", typeof(string));
            for (int i = 1; i <= 30; i++)
            {
                TotalAmountTable.Columns.Add((i * 10).ToString(), typeof(int));
            }
            TotalAmountTable.Columns.Add(">300", typeof(int));
            TotalAmountTable.Columns.Add("總計", typeof(int));


            DateTime firstDay = new DateTime(1000, 1, 1);
            int entryTimeIdx = -1;
            int departureTimeIdx = -1;
            int totalAmountIdx = -1;

            int startIndex;
            int endIndex;

            foreach (DataColumn column in oTable.Columns)
            {
                string name = string.Empty;
                int newColumnIndex = column.Ordinal;

                startIndex = 3;
                endIndex = rowCount - 1;

                //type1 停車票號,進場時間,出場時間,發票號碼,收費金額
                //type2 入場時間,繳費時間,票號／月租卡,發票號碼,金額
                string Type1_Name = oTable.Rows[2][column].ToString();
                string Type2_Name = string.Empty;

                if (Type1_Name == "停車票號")
                {
                    name = Type1_Name;
                    for (int i = startIndex; i < endIndex; i++)
                    {
                        tempTable.Rows.Add(tempTable.NewRow());
                    }
                }
                else if (Type1_Name == "進場時間")
                {
                    name = "入場時間";
                }
                else if (Type1_Name == "出場時間" || Type1_Name == "繳費時間")
                {
                    name = "出場時間";
                }
                else if (Type1_Name == "發票號碼")
                {
                    name = Type1_Name;
                }
                else if (Type1_Name == "收費金額")
                {
                    name = Type1_Name;
                }
                else
                {
                    if (!string.IsNullOrEmpty(oTable.Rows[5][column].ToString()))
                    {
                        Type2_Name = oTable.Rows[5][column].ToString();
                        startIndex = 6;
                        endIndex = rowCount - 2;
                    }
                    else if (!string.IsNullOrEmpty(oTable.Rows[6][column].ToString()))
                    {
                        Type2_Name = oTable.Rows[6][column].ToString();
                        startIndex = 7;
                        endIndex = rowCount - 1;
                    }
                    switch (Type2_Name)
                    {
                        case "入場時間":
                            name = Type2_Name;
                            for (int i = startIndex; i < endIndex; i++)
                            {
                                tempTable.Rows.Add(tempTable.NewRow());
                            }
                            break;
                        case "繳費時間":
                            name = "出場時間";
                            break;
                        case "發票號碼":
                            name = Type2_Name;
                            break;
                        case "票號／月租卡":
                            name = "停車票號";
                            break;
                        case "金額":
                            name = "收費金額";
                            break;
                        default:
                            break;
                    }
                }

                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                //Now 停車票號,入場時間,出場時間,發票號碼,收費金額
                int tempTableCount = tempTable.Rows.Count;
                for (int i = 0; i < endIndex - startIndex; i++)
                {
                    try
                    {
                        string d = oTable.Rows[i + startIndex][newColumnIndex].ToString();
                        if (name == "入場時間")
                        {
                            tempTable.Rows[i][name] = TimeReplace(d);
                        }
                        else if (name == "出場時間")
                        {
                            tempTable.Rows[i][name] = TimeReplace(d);
                            DateTime departureTime = DateTime.Parse(tempTable.Rows[i]["出場時間"].ToString());
                            DateTime entryTime = DateTime.Parse(tempTable.Rows[i]["入場時間"].ToString());
                            if (firstDay == new DateTime(1000, 1, 1))
                            {
                                firstDay = DateTime.Parse(departureTime.ToString("yyyy/MM/dd 00:00:00"));
                                DataRow dr = EntryTimeTable.NewRow();
                                dr["日期／時間"] = firstDay.ToString("yyyy/MM/dd");
                                for (int j = 1; j < EntryTimeTable.Columns.Count; j++)
                                {
                                    dr[j] = 0;
                                }
                                EntryTimeTable.Rows.Add(dr);
                                entryTimeIdx = 0;

                                dr = StayTimeTable.NewRow();
                                dr["日期／小時"] = firstDay.ToString("yyyy/MM/dd");
                                for (int j = 1; j < StayTimeTable.Columns.Count; j++)
                                {
                                    dr[j] = 0;
                                }
                                StayTimeTable.Rows.Add(dr);
                                departureTimeIdx = 0;

                                //TotalAmountTable
                                dr = TotalAmountTable.NewRow();
                                dr["日期／元"] = firstDay.ToString("yyyy/MM/dd");
                                for (int j = 1; j < TotalAmountTable.Columns.Count; j++)
                                {
                                    dr[j] = 0;
                                }
                                TotalAmountTable.Rows.Add(dr);
                                totalAmountIdx = 0;
                            }

                            //EntryTimeTable
                            double entryTimeDaysDiff = new TimeSpan(entryTime.Ticks - firstDay.Ticks).TotalDays;
                            if (entryTimeDaysDiff >= 0)
                            {
                                while ((int)entryTimeDaysDiff > entryTimeIdx)
                                {
                                    DataRow dr = EntryTimeTable.NewRow();
                                    dr["日期／時間"] = firstDay.AddDays(entryTimeIdx + 1).ToString("yyyy/MM/dd");
                                    for (int j = 1; j < EntryTimeTable.Columns.Count; j++)
                                    {
                                        dr[j] = 0;
                                    }
                                    EntryTimeTable.Rows.Add(dr);
                                    entryTimeIdx++;
                                }
                                string entryTimeHH = entryTime.Hour.ToString();
                                EntryTimeTable.Rows[(int)entryTimeDaysDiff][entryTimeHH] = Convert.ToInt32(EntryTimeTable.Rows[(int)entryTimeDaysDiff][entryTimeHH].ToString()) + 1;
                                EntryTimeTable.Rows[(int)entryTimeDaysDiff]["總計"] = Convert.ToInt32(EntryTimeTable.Rows[(int)entryTimeDaysDiff]["總計"].ToString()) + 1;
                            }

                            //StayTimeTable
                            double departTimeDaysDiff = new TimeSpan(departureTime.Ticks - firstDay.Ticks).TotalDays;
                            double stayTimehoursDiff = new TimeSpan(departureTime.Ticks - entryTime.Ticks).TotalHours;
                            int stayTimeMinutesDiff = new TimeSpan(departureTime.Ticks - entryTime.Ticks).Minutes;

                            while ((int)departTimeDaysDiff > departureTimeIdx)
                            {
                                DataRow dr = StayTimeTable.NewRow();
                                dr["日期／小時"] = firstDay.AddDays(departureTimeIdx + 1).ToString("yyyy/MM/dd");
                                for (int j = 1; j < StayTimeTable.Columns.Count; j++)
                                {
                                    dr[j] = 0;
                                }
                                StayTimeTable.Rows.Add(dr);
                                departureTimeIdx++;
                            }
                            //日期／小時,<0.5,<1,01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,>=20,總計
                            if (stayTimehoursDiff >= 20)
                            {
                                StayTimeTable.Rows[(int)departTimeDaysDiff][">=20"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff][">=20"].ToString()) + 1;
                                StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"].ToString()) + 1;
                            }
                            else if (stayTimehoursDiff >= 1)
                            {
                                StayTimeTable.Rows[(int)departTimeDaysDiff][((int)stayTimehoursDiff).ToString()] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff][((int)stayTimehoursDiff).ToString()].ToString()) + 1;
                                StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"].ToString()) + 1;
                            }
                            else if (stayTimeMinutesDiff >= 30)
                            {
                                StayTimeTable.Rows[(int)departTimeDaysDiff]["<1"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff]["<1"].ToString()) + 1;
                                StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"].ToString()) + 1;
                            }
                            else
                            {
                                StayTimeTable.Rows[(int)departTimeDaysDiff]["<0.5"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff]["<0.5"]) + 1;
                                StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Convert.ToInt32(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"].ToString()) + 1;
                            }
                        }
                        else if (name == "收費金額")
                        {
                            tempTable.Rows[i][name] = d;
                            DateTime departureTime = DateTime.Parse(tempTable.Rows[i]["出場時間"].ToString());

                            //TotalAmountTable
                            double departureTimeDaysDiff = new TimeSpan(departureTime.Ticks - firstDay.Ticks).TotalDays;
                            if (departureTimeDaysDiff >= 0)
                            {
                                while ((int)departureTimeDaysDiff > totalAmountIdx)
                                {
                                    DataRow dr = TotalAmountTable.NewRow();
                                    dr["日期／元"] = firstDay.ToString("yyyy/MM/dd");
                                    for (int j = 1; j < TotalAmountTable.Columns.Count; j++)
                                    {
                                        dr[j] = 0;
                                    }
                                    TotalAmountTable.Rows.Add(dr);
                                    totalAmountIdx++;
                                }
                                decimal totalAmount = Math.Ceiling(Convert.ToDecimal(d) / 10);
                                if (totalAmount > 30)
                                {
                                    TotalAmountTable.Rows[(int)departureTimeDaysDiff][">300"] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff][">300"].ToString()) + Convert.ToDecimal(d);
                                    TotalAmountTable.Rows[(int)departureTimeDaysDiff]["總計"] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff]["總計"].ToString()) + Convert.ToDecimal(d);
                                }
                                else
                                {
                                    TotalAmountTable.Rows[(int)departureTimeDaysDiff][(int)totalAmount] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff][(int)totalAmount].ToString()) + Convert.ToDecimal(d);
                                    TotalAmountTable.Rows[(int)departureTimeDaysDiff]["總計"] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff]["總計"].ToString()) + Convert.ToDecimal(d);
                                }
                            }
                        }
                        else
                        {
                            tempTable.Rows[i][name] = d;
                        }
                    }
                    catch (Exception ex)
                    {
                        var st = new StackTrace(ex, true);
                        // Get the top stack frame
                        var frame = st.GetFrame(0);
                        // Get the line number from the stack frame
                        var line = frame.GetFileLineNumber();
                        using (StreamWriter sw = new StreamWriter(eventFile, true))
                        {
                            sw.WriteLine("error1: " + DateTime.Now.ToString() + ",line: " + line + ", " + ex.Message);
                        }
                    }
                }
                DataTableToCSV(tempTable, csvFilePath);
                DataTableToCSV(EntryTimeTable, entryTimeFile);
                DataTableToCSV(StayTimeTable, stayTimeFile);
                DataTableToCSV(TotalAmountTable, totalAmountFile);
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

        private string TimeReplace(string time)
        {
            string dateTime = string.Empty;
            try
            {
                string tempTime = time;
                for (int i = 0; i < 24; i++)
                {
                    string HH = i.ToString("00");
                    tempTime = tempTime.Replace($"{HH}{HH}:", $"{HH}:");
                }
                dateTime = DateTime.Parse(tempTime).ToString("yyyy/MM/dd HH:mm:ss");
            }
            catch (Exception ex)
            {
                using (StreamWriter sw = new StreamWriter(eventFile, true))
                {
                    sw.WriteLine("Error: DateTime failure, " + DateTime.Now.ToString() + ", " + ex.Message);
                }
            }
            return dateTime;
        }

        private void DataTableToCSV(DataTable datatable, string FilePath)
        {
            StringBuilder sb = new StringBuilder();
            string[] columnNames = datatable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
            sb.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in datatable.Rows)
            {
                string[] fields = row.ItemArray.
                Select(field => field.ToString()).ToArray();
                sb.AppendLine(string.Join(",", fields));
            }
            System.IO.File.WriteAllText(FilePath, sb.ToString(), Encoding.GetEncoding("big5"));
        }
    }
}
