using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;

namespace ExcelToCsvApp
{
    //메인 윈도우: Window 상속받는 클래스, 필수요소
    //partial class인데 다른곳에서 기본적으로 Window를 상속받음
    public partial class MainWindow
    {
        //ObservableCollection은 문자열 리스트, WPF와 동기화되는 클래스
        private ObservableCollection<string> excelFiles = new ObservableCollection<string>();

        //클래스 생성자
        public MainWindow()
        {
            //XAML을 불러와서 초기화하는 메서드, 필수
            InitializeComponent();

            //FileListBox는 xaml에서 가져오는데, 이걸 문자열 리스트랑 연결함
            //이렇게 바인딩 해놓으면 런타임에서 추가하면 알아서 반영됨
            //FileListBox 선언 위치 따라가면 xaml 파일이 나옴. 즉, 이런거 선언은 xaml 파일에서 해야함
            FileListBox.ItemsSource = excelFiles;

            // 저장 위치 불러오기
            var savedPath = UserSettings.SaveFolderPath;
            if (!string.IsNullOrEmpty(savedPath) && Directory.Exists(savedPath))
                SavePathTextBox.Text = savedPath;
        }

        /// <summary>
        /// 이벤트 핸들러 함수. 바인딩은 xaml에서 되어있음
        /// </summary>
        private void Window_DragOver(object sender, DragEventArgs e)
        {
            //드래그된 데이터가 파일 드롭인지 체크함
            //e.Data는 IDataObject라는 인터페이스라서 데이터 체크가 필요함
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            //이벤트가 처리되었음: 중복 처리 방지인듯
            e.Handled = true;
        }

        /// <summary>
        /// 이벤트 핸들러
        /// </summary>
        private void Window_Drop(object sender, DragEventArgs e)
        {
            //드래그 드롭한게 파일이 아니다: 무시
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            //Data들에서 파일 경로 전체를 얻어오고, string[]으로 캐스팅
            var droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
            
            //Linq: .xlsx이나 .xls로 끝나는 파일만 가져옴
            //StringComparison.OrdinalIgnoreCase는 대소문자 무시
            var validFiles = droppedFiles.Where(f =>
                f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

            foreach (var file in validFiles)
            {
                //이미 들어있는 파일이 아니라면 추가
                if (!excelFiles.Contains(file))
                    excelFiles.Add(file);
            }
        }

        /// <summary>
        /// 버튼 클릭 바인딩 함수
        /// </summary>
        private void OnConvertClicked(object sender, RoutedEventArgs e)
        {
            //csv 저장 경로 가져옴
            var saveFolder = SavePathTextBox.Text.Trim();

            //저장 경로가 빈 경우 예외처리
            if (string.IsNullOrEmpty(saveFolder) || !Directory.Exists(saveFolder))
            {
                MessageBox.Show("유효한 저장 폴더 경로를 입력해주세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            //엑셀 올린 파일이 없는 경우 예외처리
            if (excelFiles.Count == 0)
            {
                MessageBox.Show("변환할 엑셀 파일을 드래그해서 추가해주세요.", "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int successCount = 0;
            var failedFileName = new StringBuilder();

            //파일마다 처리
            foreach (var excelPath in excelFiles)
            {
                try
                {
                    //파일 가져와서 엑셀로 변환하고 csv 파일로 변환해서 저장
                    string csvContent = ExcelConverter.ConvertToCsv(excelPath);
                    var fileName = Path.GetFileNameWithoutExtension(excelPath);
                    var csvPath = Path.Combine(saveFolder, fileName + ".csv");
                    File.WriteAllText(csvPath, csvContent, Encoding.UTF8);
                    successCount++;
                }
                catch (Exception ex)
                {
                    failedFileName.AppendLine($"{Path.GetFileName(excelPath)} 변환 실패: {ex.Message}");
                }
            }

            string msg = $"변환 완료!\n성공: {successCount}개";
            if (failedFileName.Length > 0)
                msg += $"\n실패:\n{failedFileName}";

            //결과 메세지 출력
            MessageBox.Show(msg, "결과", MessageBoxButton.OK, MessageBoxImage.Information);

            // 저장 경로 저장
            UserSettings.SaveFolderPath = saveFolder;
        }
    }
}
