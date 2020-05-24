using MinoriEditorShell.Services;
using MvvmCross.Commands;
using MvvmCross.Logging;
using MvvmCross.Navigation;
using MvvmCross.ViewModels;

namespace VisualBasicUpgradeAssistant.Core.ViewModels
{
    public class MainViewModel : MvxNavigationViewModel
    {
        private readonly IMesDocumentManager _documentManager;
        private readonly IMvxNavigationService _navigationService;

        public MainViewModel(
            IMesDocumentManager documentManager,
            IMvxLogProvider logProvider,
            IMesWindow mesWindow,
            IMvxNavigationService navigationService
            ) : base(logProvider, navigationService)
        {
            _documentManager = documentManager;
            _navigationService = navigationService;

            mesWindow.DisplayName = "Visual Basic Upgrade Assistant (Alpha)";
        }

        public IMvxCommand OpenOmegaCommand => new MvxCommand(() => NavigationService.Navigate<OmegaViewModel>());
    }
}
