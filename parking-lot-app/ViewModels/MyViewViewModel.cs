using Prism.Commands;
using Prism.Mvvm;
using Prism.Regions;
using System.Threading.Tasks;
using parking_lot_app.Views;
using parking_lot_app.Model.MyView;

namespace parking_lot_app.ViewModels
{
    public class MyViewModel : BindableBase , INavigationAware
    {
        private string message;
        private readonly IRegionManager regionManager;

        public string Message
        {
            get { return message; }
            set { SetProperty(ref message, value); }
        }

        public DelegateCommand GoNextCommand { get; set; }
        public DelegateCommand OpenFile { get; set; }
        
        public int Counter { get; set; }
        public MyViewModel(IRegionManager regionManager)
        {
            File f1 = new File(@"C:\Users\Ryan\Projects\parking-lot-app\富軥1029-30.xls");
            this.regionManager = regionManager;
            OpenFile  = new DelegateCommand(() =>
            {
                f1.OpenFile();
            });
            GoNextCommand = new DelegateCommand(() =>
            {
                regionManager.RequestNavigate("ContentRegion", nameof(View1));
            });
            //File f2 = new File(@"C:\Users\Ryan\Projects\parking-lot-app\慈祥10月.xls");
            //File f3 = new File(@"C:\Users\Ryan\Projects\parking-lot-app\豐俊1029.30.xls");
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
            Message = navigationContext.NavigationService.Journal.CanGoBack == false ? "尚未開始導航 "+ Counter++ : "可以回上一頁";
            //regionManager.Regions["ContentRegion"].NavigationService.Journal.Clear();
        }
    }
}
