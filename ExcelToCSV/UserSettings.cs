using System;
using System.IO;

namespace ExcelToCsvApp
{
    public static class UserSettings
    {
        private static readonly string ConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelToCsvApp",
            "config.txt");

        public static string SaveFolderPath
        {
            get
            {
                try
                {
                    if (!File.Exists(ConfigFilePath))
                        return string.Empty;
                    return File.ReadAllText(ConfigFilePath);
                }
                catch
                {
                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    var dir = Path.GetDirectoryName(ConfigFilePath);
                    if (!Directory.Exists(dir))
                        Directory.CreateDirectory(dir ?? throw new InvalidOperationException());
                    File.WriteAllText(ConfigFilePath, value);
                }
                catch
                {
                    // ignored
                }
            }
        }
    }
}