using MvvmCross.Logging;
using MvvmCross.Navigation;
using MvvmCross.ViewModels;

namespace VisualBasicUpgradeAssistant.Core.ViewModels
{
    public class MainViewModel : MvxNavigationViewModel
    {
        private readonly IMvxNavigationService _navigationService;

        public MainViewModel(
            IMvxLogProvider logProvider,
            IMvxNavigationService navigationService
            ) : base(logProvider, navigationService)
        {
            _navigationService = navigationService;
        }
    }
}
