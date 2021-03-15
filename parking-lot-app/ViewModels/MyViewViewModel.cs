using Prism.Commands;
using Prism.Mvvm;
using Prism.Regions;
using System.Threading.Tasks;
using parking_lot_app.Views;
using parking_lot_app.Model.MyView;
using LiveCharts;
using LiveCharts.Wpf;
using System.Windows;
using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;

namespace parking_lot_app.ViewModels
{
    public class MyViewModel : BindableBase, INavigationAware
    {
        private string message;
        private string log;

        private SeriesCollection entryTimeFileSeriesCollection;
        private SeriesCollection stayTimeFileSeriesCollection;
        private SeriesCollection totalAmountFileSeriesCollection;

        private List<string> entryTimeLabels;
        private List<string> stayTimeLabels;
        private List<string> totalAmountLabels;

        private string spaceValue;
        private string ceilingValue;
        private string floorValue;


        private readonly IRegionManager regionManager;

        public string Message
        {
            get { return message; }
            set { SetProperty(ref message, value); }
        }

        public string Log
        {
            get { return log; }
            set { SetProperty(ref log, value); }
        }

        public SeriesCollection EntryTimeFileSeriesCollection
        {
            get { return entryTimeFileSeriesCollection; }
            set { SetProperty(ref entryTimeFileSeriesCollection, value); }
        }

        public SeriesCollection StayTimeFileSeriesCollection
        {
            get { return stayTimeFileSeriesCollection; }
            set { SetProperty(ref stayTimeFileSeriesCollection, value); }
        }

        public SeriesCollection TotalAmountFileSeriesCollection
        {
            get { return totalAmountFileSeriesCollection; }
            set { SetProperty(ref totalAmountFileSeriesCollection, value); }
        }

        public List<string> EntryTimeLabels
        {
            get { return entryTimeLabels; }
            set { SetProperty(ref entryTimeLabels, value); }
        }
        public List<string> StayTimeLabels
        {
            get { return stayTimeLabels; }
            set { SetProperty(ref stayTimeLabels, value); }
        }
        public List<string> TotalAmountLabels
        {
            get { return totalAmountLabels; }
            set { SetProperty(ref totalAmountLabels, value); }
        }

        public string SpaceValue
        {
            get { return spaceValue; }
            set { SetProperty(ref spaceValue, value); }
        }
        public string CeilingValue
        {
            get { return ceilingValue; }
            set { SetProperty(ref ceilingValue, value); }
        }
        public string FloorValue
        {
            get { return floorValue; }
            set { SetProperty(ref floorValue, value); }
        }

        public DelegateCommand GoNextCommand { get; set; }
        public DelegateCommand OpenFile { get; set; }

        public int Counter { get; set; }
        public MyViewModel(IRegionManager regionManager)
        {

            IniApi ini = new IniApi("./Setting.ini");
            InitIniData(ini);

            //mylineseries.Stroke = System.Windows.Media.Brushes.Black;
            //mylineseries.StrokeThickness = 10;
            //mylineseries.StrokeDashArray = new System.Windows.Media.DoubleCollection { 2 };
            //mylineseries.LineSmoothness = 1;
            //mylineseries.Fill = System.Windows.Media.Brushes.LightBlue;

            ColumnSeries EntryTimeSeries = new ColumnSeries();
            ColumnSeries StayTimeSeries = new ColumnSeries();
            LineSeries TotalAmountSeries = new LineSeries();

            EntryTimeSeries.Title = "入場表";
            StayTimeSeries.Title = " 停車表";
            TotalAmountSeries.Title = "金額表";

            EntryTimeSeries.DataLabels = true;
            StayTimeSeries.DataLabels = true;
            TotalAmountSeries.DataLabels = true;

            //EntryTimeSeries.Stroke = System.Windows.Media.Brushes.Black;
            //StayTimeSeries.Stroke = System.Windows.Media.Brushes.Red;
            //TotalAmountSeries.Stroke = System.Windows.Media.Brushes.Blue;

            TotalAmountSeries.LineSmoothness = 0;
            TotalAmountSeries.PointGeometry = null;
            //Labels = new List<string> { "1", "3", "2", "4", "-3", "5", "2", "1" };
            ////添加折线图的数据
            //mylineseries.Values = new ChartValues<double>(temp);
            //SeriesCollection = new SeriesCollection { };
            //SeriesCollection.Add(mylineseries);

            EntryTimeFileSeriesCollection = new SeriesCollection(new ColumnSeries());
            StayTimeFileSeriesCollection = new SeriesCollection(new ColumnSeries());
            TotalAmountFileSeriesCollection = new SeriesCollection(new LineSeries());

            try
            {
                File f1 = new File();
                OpenFile = new DelegateCommand(() =>
                {
                    WriteIniData(ini);
                    DataTable[] tableList = f1.OpenFile(SpaceValue, CeilingValue, FloorValue);
                    if (tableList == null)
                    {
                        return;
                    }

                    List<decimal> entryTimeValues = ConvertTableToList(tableList[0]);
                    List<decimal> stayTimeValues = ConvertTableToList(tableList[1]);
                    List<decimal> totalAmountValues = ConvertTableToList(tableList[2]);

                    EntryTimeLabels = ConvertTableToAxisX(tableList[0]);
                    StayTimeLabels = ConvertTableToAxisX(tableList[1]);
                    TotalAmountLabels = ConvertTableToAxisX(tableList[2]);

                    EntryTimeSeries.Values = new ChartValues<decimal>(entryTimeValues);
                    EntryTimeFileSeriesCollection.Clear();
                    EntryTimeFileSeriesCollection.Add(EntryTimeSeries);

                    StayTimeSeries.Values = new ChartValues<decimal>(stayTimeValues);
                    StayTimeFileSeriesCollection.Clear();
                    StayTimeFileSeriesCollection.Add(StayTimeSeries);

                    TotalAmountSeries.Values = new ChartValues<decimal>(totalAmountValues);
                    TotalAmountFileSeriesCollection.Clear();
                    TotalAmountFileSeriesCollection.Add(TotalAmountSeries);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("File Open Error: search for File Model, " + ex.Message);
            }
            this.regionManager = regionManager;
            GoNextCommand = new DelegateCommand(() =>
            {
                regionManager.RequestNavigate("ContentRegion", nameof(View1));
            });
        }
        private List<string> ConvertTableToAxisX(DataTable table)
        {
            List<string> entryTimeValues = new List<string> { };
            for (int i = 1; i < table.Columns.Count - 1; i++)
            {
                var t = table.Columns[i].ColumnName;
                entryTimeValues.Add(t);
            }
            return entryTimeValues;
        }
        private List<decimal> ConvertTableToList(DataTable table)
        {
            List<decimal> entryTimeValues = new List<decimal> { };
            for (int i = 1; i < table.Columns.Count - 1; i++)
            {
                Console.WriteLine("III: " + table.Rows[table.Rows.Count - 1][i].ToString());
                var t = table.Rows[table.Rows.Count - 1][i].ToString();
                decimal.TryParse(t, out decimal tt);
                entryTimeValues.Add(tt);
            }
            return entryTimeValues;
        }
        private void InitIniData(IniApi ini)
        {
            ReadIniData(ini);
            WriteIniData(ini);
        }

        private void ReadIniData(IniApi ini)
        {
            SpaceValue = ini.ReadIniFile("Setting", "space_value", @"10");
            CeilingValue = ini.ReadIniFile("Setting", "floor_value", @"10");
            FloorValue = ini.ReadIniFile("Setting", "ceiling_value", @"150");
        }

        private void WriteIniData(IniApi ini)
        {
            ini.WriteIniFile("Setting", "space_value", SpaceValue);
            ini.WriteIniFile("Setting", "floor_value", CeilingValue);
            ini.WriteIniFile("Setting", "ceiling_value", FloorValue);
        }

        public bool IsNavigationTarget(NavigationContext navigationContext)
        {
            return true;
        }

        public void OnNavigatedFrom(NavigationContext navigationContext)
        {
        }

        public async void OnNavigatedTo(NavigationContext navigationContext)
        {
            await Task.Yield();
            Message = navigationContext.NavigationService.Journal.CanGoBack == false ? "尚未開始導航 " + Counter++ : "可以回上一頁";
        }
    }
}
