using MvvmCross.IoC;
using MvvmCross.ViewModels;
using VisualBasicUpgradeAssistant.Core.ViewModels;

namespace VisualBasicUpgradeAssistant.Core
{
    public class App : MvxApplication
    {
        public override void Initialize()
        {
            CreatableTypes()
                .EndingWith("Service")
                .AsInterfaces()
                .RegisterAsLazySingleton();

            CreatableTypes()
                .EndingWith("ViewModel")
                .AsTypes()
                .RegisterAsDynamic();

            RegisterAppStart<MainViewModel>();
        }
    }
}
