using System;
using System.IO;

namespace ExcelToCsvApp
{
    public static class UserSettings
    {
        /// <summary>
        /// 파일 경로 string
        /// </summary>
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
                    //경로 저장 파일이 있는지 검사하고 return
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