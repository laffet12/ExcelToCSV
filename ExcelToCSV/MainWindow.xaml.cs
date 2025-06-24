using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;

namespace ExcelToCsvApp
{
    public partial class MainWindow
    {
        private ObservableCollection<string> excelFiles = new ObservableCollection<string>();

        public MainWindow()
        {
            InitializeComponent();

            FileListBox.ItemsSource = excelFiles;

            // 저장 위치 불러오기
            var savedPath = UserSettings.SaveFolderPath;
            if (!string.IsNullOrEmpty(savedPath) && Directory.Exists(savedPath))
                SavePathTextBox.Text = savedPath;
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            var droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
            var validFiles = droppedFiles.Where(f =>
                f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

            foreach (var file in validFiles)
            {
                if (!excelFiles.Contains(file))
                    excelFiles.Add(file);
            }
        }

        private void OnConvertClicked(object sender, RoutedEventArgs e)
        {
            var saveFolder = SavePathTextBox.Text.Trim();

            if (string.IsNullOrEmpty(saveFolder) || !Directory.Exists(saveFolder))
            {
                MessageBox.Show("유효한 저장 폴더 경로를 입력해주세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (excelFiles.Count == 0)
            {
                MessageBox.Show("변환할 엑셀 파일을 드래그해서 추가해주세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int successCount = 0;
            var failedFiles = new StringBuilder();

            foreach (var excelPath in excelFiles)
            {
                try
                {
                    string csvContent = ExcelConverter.ConvertToCsv(excelPath);
                    var fileName = Path.GetFileNameWithoutExtension(excelPath);
                    var csvPath = Path.Combine(saveFolder, fileName + ".csv");
                    File.WriteAllText(csvPath, csvContent, Encoding.UTF8);
                    successCount++;
                }
                catch (Exception ex)
                {
                    failedFiles.AppendLine($"{Path.GetFileName(excelPath)} 변환 실패: {ex.Message}");
                }
            }

            string msg = $"변환 완료!\n성공: {successCount}개";
            if (failedFiles.Length > 0)
                msg += $"\n실패:\n{failedFiles}";

            MessageBox.Show(msg, "결과", MessageBoxButton.OK, MessageBoxImage.Information);

            // 저장 경로 저장
            UserSettings.SaveFolderPath = saveFolder;
        }
    }
}
