using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;

using IPConfig.Properties;

using HcLang = HandyControl.Properties.Langs.Lang;

namespace IPConfig.Languages;

/// <summary>
/// <see href="https://www.codinginfinity.me/posts/localization-of-a-wpf-app-the-simple-approach/">localization-of-a-wpf-app-the-simple-approach</see>
/// </summary>
public sealed class LangSource : INotifyPropertyChanged
{
    #region Singleton

    public static LangSource Instance { get; } = new LangSource();

    static LangSource()
    {
        Instance.SetLanguage(String.IsNullOrWhiteSpace(Settings.Default.Language) ? "en" : Settings.Default.Language);
    }

    private LangSource()
    { }

    #endregion Singleton

    private CultureInfo _currentCulture = CultureInfo.CurrentUICulture;

    public CultureInfo CurrentCulture
    {
        get => _currentCulture;
        set
        {
            if (!_currentCulture.Equals(value))
            {
                _currentCulture = value;
                PropertyChanged?.Invoke(this, new(String.Empty));
                LanguageChanged?.Invoke(this, new(nameof(CurrentCulture)));
            }
        }
    }

    public string? this[string key] => Lang.ResourceManager.GetString(key, CurrentCulture);

    public string this[LangKey key] => Lang.ResourceManager.GetString(key.ToString(), CurrentCulture)!;

    public event PropertyChangedEventHandler? LanguageChanged;

    public event PropertyChangedEventHandler? PropertyChanged;

    public static List<CultureInfo> GetAvailableCultures()
    {
        return new() { CultureInfo.GetCultureInfo("en") };
    }

    public void SetLanguage(string locale)
    {
        if (String.IsNullOrWhiteSpace(locale))
        {
            locale = "en";
        }

        var newCulture = CultureInfo.GetCultureInfo(locale);
        Lang.Culture = newCulture;
        CurrentCulture = newCulture;
        HcLang.Culture = newCulture;

        Settings.Default.Language = locale;
        Settings.Default.Save();
    }
}
