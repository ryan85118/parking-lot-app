using Prism.Commands;
using Prism.Mvvm;
using Prism.Regions;
using System.Threading.Tasks;
using parking_lot_app.Views;
using parking_lot_app.Model.MyView;
using System.Windows;
using System;

namespace parking_lot_app.ViewModels
{
    public class MyViewModel : BindableBase, INavigationAware
    {
        private string message;
        private string log;

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

        public DelegateCommand GoNextCommand { get; set; }
        public DelegateCommand OpenFile { get; set; }

        public int Counter { get; set; }
        public MyViewModel(IRegionManager regionManager)
        {
            try
            {
                File f1 = new File();
                OpenFile = new DelegateCommand(() =>
                {
                    f1.OpenFile();
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
