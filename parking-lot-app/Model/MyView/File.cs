using Microsoft.Win32;
using System;
using System.Collections.Specialized;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace parking_lot_app.Model.MyView
{
    internal class File
    {
        private readonly IniApi ini = new IniApi("./Setting.ini");
        private int space_value;
        private int floor_value;
        private int ceiling_value;

        private static readonly string eventFilePath = "./event";
        private static readonly string eventFileName = "event.txt";
        private static readonly string eventFile = Path.Combine(eventFilePath, eventFileName);
        private static readonly string outputFilePath = "../output";
        private static readonly string entryTimeFilePath = "../output/單月入場";
        private static readonly string stayTimeFilePath = "../output/單月停車";
        private static readonly string totalAmountFilePath = "../output/單月金額";
        private string entryTimeFile;
        private string stayTimeFile;
        private string totalAmountFile;

        private string tempPath = null;
        private string tempName = null;

        public File()
        {
            Directory.CreateDirectory(eventFilePath);
        }

        public DataTable[] OpenFile(string spaceValue, string floorValue, string ceilingValue)
        {
            InitValues(spaceValue, ceilingValue, floorValue);
            OpenFileDialog odXls = new OpenFileDialog();
            //指定相應的開啟文件的目錄 AppDomain.CurrentDomain.BaseDirectory定位到Debug目錄，再根據實際情況進行目錄調整
            string folderPatha = AppDomain.CurrentDomain.BaseDirectory + @"databackup\";
            odXls.InitialDirectory = folderPatha;
            odXls.Filter = "所有 Excel 檔案(*.xl*)|*.xl*|All files (*.*)|*.*";
            odXls.RestoreDirectory = true;
            odXls.Multiselect = true;
            if (odXls.ShowDialog() == true)
            {
                foreach (string fileName in odXls.FileNames)
                {
                    using (StreamWriter sw = new StreamWriter(eventFile, true))
                    {
                        sw.WriteLine(fileName);
                        tempPath = fileName;
                        tempName = Path.GetFileNameWithoutExtension(tempPath);
                        sw.WriteLine(tempName);
                    }
                    return ReadFile();
                }
            }
            return null;
        }

        private DataTable[] ReadFile()
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
                return SaveToCSV(dt, tempFilePath);
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
            return null;
        }

        private DataTable[] SaveToCSV(DataTable oTable, string csvFilePath)
        {
            CreateDirectory();
            try
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
                    EntryTimeTable.Columns.Add(i.ToString() + "-" + (i + 1).ToString() + "點", typeof(string));
                }
                EntryTimeTable.Columns.Add("總計", typeof(int));

                //StayTimeTable
                //日期／H <0.5 <1H 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 >=24 總計
                StayTimeTable.Columns.Add("日期／H", typeof(string));
                StayTimeTable.Columns.Add("<0.5H", typeof(double));
                StayTimeTable.Columns.Add("<1H", typeof(double));
                for (int i = 1; i < 24; i++)
                {
                    StayTimeTable.Columns.Add($"{i}-{i + 1}H", typeof(double));
                }
                StayTimeTable.Columns.Add(">=24H", typeof(double));
                StayTimeTable.Columns.Add("總計", typeof(double));

                //TotalAmountTable
                TotalAmountTable.Columns.Add("日期／元", typeof(string));
                int floorIndex = floor_value / space_value;
                int ceilingIndex = ceiling_value / space_value;

                for (int i = 0; i <= ceilingIndex - floorIndex; i++)
                {
                    TotalAmountTable.Columns.Add((i * space_value + floor_value).ToString(), typeof(int));
                }
                TotalAmountTable.Columns.Add(">" + ceiling_value, typeof(int));
                TotalAmountTable.Columns.Add("總計", typeof(int));

                DateTime firstDay = new DateTime(3000, 1, 1);
                int entryTimeIdx = -1;
                int departureTimeIdx = -1;
                int totalAmountIdx = -1;

                int startIndex;
                int endIndex;
                try
                {
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
                            else if (!string.IsNullOrEmpty(oTable.Rows[7][column].ToString()))
                            {
                                Type2_Name = oTable.Rows[7][column].ToString();
                                startIndex = 8;
                                endIndex = rowCount;
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
                                case "票號":
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

                        if (name == "出場時間")
                        {
                            try
                            {
                                for (int i = 0; i < endIndex - startIndex - 1; i++)
                                {
                                    string d = oTable.Rows[i + startIndex][newColumnIndex].ToString();
                                    DateTime departureTime;
                                    try
                                    {
                                        departureTime = DateTime.Parse(TimeReplace(d));
                                    }
                                    catch
                                    {
                                        departureTime = new DateTime(3000, 1, 1);
                                    }

                                    if (firstDay > DateTime.Parse(departureTime.ToString("yyyy/MM/dd 00:00:00")))
                                    {
                                        firstDay = DateTime.Parse(departureTime.ToString("yyyy/MM/dd 00:00:00"));
                                    }
                                }
                                DataRow dr = EntryTimeTable.NewRow();
                                dr["日期／時間"] = firstDay.ToString("yyyy/MM/dd");
                                for (int j = 1; j < EntryTimeTable.Columns.Count; j++)
                                {
                                    dr[j] = 0;
                                }
                                EntryTimeTable.Rows.Add(dr);
                                entryTimeIdx = 0;

                                dr = StayTimeTable.NewRow();
                                dr["日期／H"] = firstDay.ToString("yyyy/MM/dd");
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
                            catch (Exception ex)
                            {
                                var st = new StackTrace(ex, true);
                                // Get the top stack frame
                                var frame = st.GetFrame(0);
                                // Get the line number from the stack frame
                                var line = frame.GetFileLineNumber();
                                MessageBox.Show("line" + line + "," + ex.Message);
                            }
                        }

                        //Now 停車票號,入場時間,出場時間,發票號碼,收費金額
                        int tempTableCount = tempTable.Rows.Count;
                        for (int i = 0; i < endIndex - startIndex; i++)
                        {
                            try
                            {
                                string d = oTable.Rows[i + startIndex][newColumnIndex].ToString();
                                if (string.IsNullOrEmpty(d))
                                {
                                    continue;
                                }
                                else if (name == "入場時間")
                                {
                                    tempTable.Rows[i][name] = TimeReplace(d);
                                }
                                else if (name == "出場時間")
                                {
                                    tempTable.Rows[i][name] = TimeReplace(d);
                                    DateTime departureTime = DateTime.Parse(tempTable.Rows[i]["出場時間"].ToString());

                                    DateTime entryTime = DateTime.Parse(tempTable.Rows[i]["入場時間"].ToString());

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
                                        string entryTimeHH = entryTime.Hour.ToString() + "-" + (entryTime.Hour + 1).ToString() + "點";
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
                                        dr["日期／H"] = firstDay.AddDays(departureTimeIdx + 1).ToString("yyyy/MM/dd");
                                        for (int j = 1; j < StayTimeTable.Columns.Count; j++)
                                        {
                                            dr[j] = 0;
                                        }
                                        StayTimeTable.Rows.Add(dr);
                                        departureTimeIdx++;
                                    }

                                    //日期／H,<0.5,<1,01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23,>=24,總計
                                    if (stayTimehoursDiff >= 24)
                                    {
                                        StayTimeTable.Rows[(int)departTimeDaysDiff][">=24H"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff][">=24H"]) + 1, 2);
                                        StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"]) + 1, 2);
                                    }
                                    else if (stayTimehoursDiff >= 1)
                                    {
                                        int stayIndex = (int)stayTimehoursDiff;

                                        StayTimeTable.Rows[(int)departTimeDaysDiff][$"{stayIndex}-{stayIndex + 1}H"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff][$"{stayIndex}-{stayIndex + 1}H"]) + 1, 2);
                                        StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"]) + 1, 2);
                                    }
                                    else if (stayTimeMinutesDiff >= 30)
                                    {
                                        StayTimeTable.Rows[(int)departTimeDaysDiff]["<1H"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff]["<1H"]) + 1, 2);
                                        StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"]) + 1, 2);
                                    }
                                    else
                                    {
                                        StayTimeTable.Rows[(int)departTimeDaysDiff]["<0.5H"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff]["<0.5H"]) + 1, 2);
                                        StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"] = Math.Round(Convert.ToDouble(StayTimeTable.Rows[(int)departTimeDaysDiff]["總計"]) + 1, 2);
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
                                        //space_value floor_value ceiling_value
                                        while ((int)departureTimeDaysDiff > totalAmountIdx)
                                        {
                                            DataRow dr = TotalAmountTable.NewRow();
                                            dr["日期／元"] = firstDay.AddDays(totalAmountIdx + 1).ToString("yyyy/MM/dd");
                                            for (int j = 1; j < TotalAmountTable.Columns.Count; j++)
                                            {
                                                dr[j] = 0;
                                            }
                                            TotalAmountTable.Rows.Add(dr);
                                            totalAmountIdx++;
                                        }
                                        int totalAmount = Convert.ToInt32(d);
                                        if (totalAmount > ceiling_value)
                                        {
                                            TotalAmountTable.Rows[(int)departureTimeDaysDiff][">" + ceiling_value] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff][">" + ceiling_value].ToString()) + Convert.ToDecimal(d);
                                            TotalAmountTable.Rows[(int)departureTimeDaysDiff]["總計"] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff]["總計"].ToString()) + Convert.ToDecimal(d);
                                        }
                                        else
                                        {
                                            int lowNumber = totalAmount / space_value - floorIndex + 1;
                                            lowNumber = lowNumber < 1 ? 1 : lowNumber;
                                            TotalAmountTable.Rows[(int)departureTimeDaysDiff][lowNumber] = Convert.ToInt32(TotalAmountTable.Rows[(int)departureTimeDaysDiff][lowNumber].ToString()) + Convert.ToDecimal(d);
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
                                StackTrace st = new StackTrace(ex, true);
                                StackFrame frame = st.GetFrame(st.FrameCount - 1);
                                using (StreamWriter sw = new StreamWriter(eventFile, true))
                                {
                                    sw.WriteLine("error1: " + DateTime.Now.ToString() + ",line: " + frame.GetFileLineNumber() + ", " + ex.Message);
                                }
                            }
                        }

                        if (name == "收費金額")
                        {
                            //EntryTimeTable
                            DataRow drr = EntryTimeTable.NewRow();
                            for (int i = 1; i < EntryTimeTable.Columns.Count; i++)
                            {
                                drr[i] = 0;
                                for (int j = 0; j < EntryTimeTable.Rows.Count; j++)
                                {
                                    drr[i] = Convert.ToInt32(drr[i]) + Convert.ToInt32(EntryTimeTable.Rows[j][i]);
                                }
                            }
                            drr["日期／時間"] = "總計";
                            EntryTimeTable.Rows.Add(drr);

                            //StayTimeTable
                            drr = StayTimeTable.NewRow();
                            for (int i = 1; i < StayTimeTable.Columns.Count; i++)
                            {
                                drr[i] = 0;
                                for (int j = 0; j < StayTimeTable.Rows.Count; j++)
                                {
                                    drr[i] = Math.Round(Convert.ToDouble(drr[i]) + Convert.ToDouble(StayTimeTable.Rows[j][i]), 2);
                                }
                            }
                            drr["日期／H"] = "總計";
                            StayTimeTable.Rows.Add(drr);

                            //TotalAmountTable
                            drr = TotalAmountTable.NewRow();
                            for (int i = 1; i < TotalAmountTable.Columns.Count; i++)
                            {
                                drr[i] = 0;
                                for (int j = 0; j < TotalAmountTable.Rows.Count; j++)
                                {
                                    drr[i] = Convert.ToInt32(drr[i]) + Convert.ToInt32(TotalAmountTable.Rows[j][i]);
                                }
                            }
                            drr["日期／元"] = "總計";
                            TotalAmountTable.Rows.Add(drr);
                        }
                    }
                    DataTableToCSV(tempTable, csvFilePath);
                    DataTableToCSV(EntryTimeTable, entryTimeFile);
                    DataTableToCSV(StayTimeTable, stayTimeFile);
                    DataTableToCSV(TotalAmountTable, totalAmountFile);
                    return new DataTable[3] { EntryTimeTable, StayTimeTable, TotalAmountTable };
                }
                catch (Exception ex)
                {
                    var st = new StackTrace(ex, true);
                    // Get the top stack frame
                    var frame = st.GetFrame(0);
                    // Get the line number from the stack frame
                    var line = frame.GetFileLineNumber();
                    MessageBox.Show("line" + line + "," + ex.Message);
                }
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                MessageBox.Show("line" + line + "," + ex.Message);
            }

            return null;
        }

        private void CreateDirectory()
        {
            Directory.CreateDirectory(eventFilePath);
            Directory.CreateDirectory(entryTimeFilePath);
            Directory.CreateDirectory(stayTimeFilePath);
            Directory.CreateDirectory(totalAmountFilePath);
        }

        private void InitValues(string spaceValue, string ceilingValue, string floorValue)
        {
            space_value = Convert.ToInt32(spaceValue);
            floor_value = Convert.ToInt32(floorValue);
            ceiling_value = Convert.ToInt32(ceilingValue);
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

                //using (StreamWriter sw = new StreamWriter(eventFile, true))
                //{
                //    sw.WriteLine("tempTime: " + tempTime);
                //}

                if (string.IsNullOrEmpty(time))
                {
                    return null;
                }
                else if (time.Contains('/'))
                {
                    dateTime = DateTime.ParseExact(time, "d/M/yyyy H:mm:ss", CultureInfo.InvariantCulture).ToString("yyyy/MM/dd HH:mm:ss");
                }
                else
                {
                    dateTime = DateTime.ParseExact(time, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture).ToString("yyyy/MM/dd HH:mm:ss");
                }

                //using (StreamWriter sw = new StreamWriter(eventFile, true))
                //{
                //    sw.WriteLine("dateTime: " + dateTime);
                //}
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