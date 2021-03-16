using parking_lot_app.Views;
using Prism.Commands;
using Prism.Mvvm;
using Prism.Regions;
using System.Threading.Tasks;

namespace parking_lot_app.ViewModels
{
    public class View1ViewModel : BindableBase, INavigationAware
    {
        private string message;
        private readonly IRegionManager regionManager;
        private readonly IRegionNavigationService regionNavigationService;

        public string Message
        {
            get { return message; }
            set { SetProperty(ref message, value); }
        }

        public int Counter { get; set; }
        public DelegateCommand GoNextCommand { get; set; }
        public DelegateCommand GoPrevCommand { get; set; }

        public View1ViewModel(IRegionManager regionManager, IRegionNavigationService regionNavigationService)
        {
            this.regionManager = regionManager;
            this.regionNavigationService = regionNavigationService;
            GoNextCommand = new DelegateCommand(() =>
            {
                regionManager.RequestNavigate("ContentRegion", nameof(PrismUserControl1));
            });
            GoPrevCommand = new DelegateCommand(() =>
            {
                regionManager.Regions["ContentRegion"].NavigationService.Journal.GoBack();
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
            Message = navigationContext.NavigationService.Journal.CanGoBack == false ? "尚未開始導航 " : "可以回上一頁 " + Counter++;
        }
    }
}