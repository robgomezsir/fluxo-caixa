using System;
using Microsoft.Win32;

namespace fluxo_caixa_crud.Services
{
    internal static class RibbonThemeHelper
    {
        public static bool IsDarkOfficeTheme()
        {
            try
            {
                string major = Globals.ThisWorkbook.Application.Version.Split('.')[0];
                using (var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{major}.0\Common"))
                {
                    string theme = key?.GetValue("UI Theme") as string;
                    if (string.IsNullOrWhiteSpace(theme) || theme.Equals("Use system setting", StringComparison.OrdinalIgnoreCase))
                        return !IsWindowsAppsLightTheme();
                    return theme.Equals("Dark Gray", StringComparison.OrdinalIgnoreCase)
                        || theme.Equals("Black", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool IsWindowsAppsLightTheme()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    object value = key?.GetValue("AppsUseLightTheme");
                    if (value is int i) return i != 0;
                }
            }
            catch { }
            return true;
        }
    }
}
