using MvvmCross.Navigation;
using MvvmCross.ViewModels;

namespace VisualBasicUpgradeAssistant.Core.ViewModels
{
    public class MainViewModel : MvxViewModel
    {
        private readonly IMvxNavigationService _navigationService;

        public MainViewModel(IMvxNavigationService navigationService)
        {
            _navigationService = navigationService;
        }
    }
}
