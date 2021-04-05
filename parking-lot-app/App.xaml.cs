using parking_lot_app.Views;
using Prism.Ioc;
using Prism.Regions;
using Prism.Unity;
using System.Windows;

namespace parking_lot_app
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : PrismApplication
    {
        protected override Window CreateShell()
        {
            var w = Container.Resolve<MainWindow>();
            return w;
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterForNavigation<MyView>();
            //containerRegistry.RegisterForNavigation<object,MyView>(nameof(MyView));
        }

        // 在這裡指定 Region 要顯示的 View ，也是可行的
        protected override void Initialize()
        {
            base.Initialize();
            //IContainerProvider container = (App.Current as PrismApplication).Container;
            IContainerProvider container = Container;
            IRegionManager regionManager = container.Resolve<IRegionManager>();
            //regionManager.RegisterViewWithRegion("ContentRegion", typeof(MyView));

            //var view = container.Resolve<MyView>();
            regionManager.RequestNavigate("ContentRegion", nameof(MyView));

            //IRegion region = regionManager.Regions["ContentRegion"];
            //var view = container.Resolve<MyView>();
            //region.Add(view);
        }
    }
}