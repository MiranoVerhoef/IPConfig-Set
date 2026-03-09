using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

using HandyControl.Data;

using IPConfig.Helpers;
using IPConfig.Properties;

namespace IPConfig.ViewModels;

public partial class ThemeSwitchButtonViewModel : ObservableObject
{
    #region Relay Commands

    [RelayCommand]
    private static void ChangeTheme(SkinType? skin)
    {
        ThemeManager.UpdateSkin(SkinType.Default);

        Settings.Default.Theme = SkinType.Default.ToString();
        Settings.Default.Save();
    }

    [RelayCommand]
    private static void Loaded()
    {
        ChangeTheme(SkinType.Default);
    }

    #endregion Relay Commands
}
