using Prism.Mvvm;

namespace parking_lot_app.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        private string _title = "自動化圖表用程式";

        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        public MainWindowViewModel()
        {
        }
    }
}