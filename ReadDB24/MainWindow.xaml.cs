using ReadDB24.Models;
using ReadDB24.Services;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ReadDB24
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SqliteDBService sqliteDBService;
        private readonly string dataBasePath = $"{Environment.CurrentDirectory}\\24.db";

        public MainWindow()
        {
            InitializeComponent();
            //Uri iconUri = new Uri($"{Environment.CurrentDirectory}\\Icons\\database.png", UriKind.RelativeOrAbsolute);
            //this.Icon = BitmapFrame.Create(iconUri);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            fullNameText.Text = "Петрик П'яточкін";
            startDateText.Text = DateTime.Now.AddDays(-10).ToShortDateString();
            endDateText.Text = DateTime.Now.ToShortDateString();
            statusTextLabel.Foreground = Brushes.Blue;
        }

        private async void runButton_Click(object sender, RoutedEventArgs e)
        {
            if (DateTime.Parse(startDateText.Text) > DateTime.Parse(endDateText.Text))
            {
                statusTextLabel.Text = "Початкова дата не може бути більшою за кінцеву дату.";
                statusTextLabel.Foreground = Brushes.Red;
                return;
            }

            if (fullNameText.Text.Trim().Length < 0)
            {
                statusTextLabel.Text = "П.І.Б. повинно містити мінімум 3 символи.";
                statusTextLabel.Foreground = Brushes.Red;
                return;
            }

            statusTextLabel.Foreground = Brushes.Blue;
            statusTextLabel.Text = "Почекайте будь-ласка...";

            var sw = new Stopwatch();
            
            sw.Start();
            sqliteDBService = new SqliteDBService(dataBasePath);
            var tableDataList = await sqliteDBService.GetRecordsByName(fullNameText.Text.Replace("'", "''"), DateTime.Parse(startDateText.Text).ToString("yyyy-MM-dd"), DateTime.Parse(endDateText.Text).ToString("yyyy-MM-dd"));
            sw.Stop();
            
            statusTextLabel.Foreground = Brushes.Blue;
            if (tableDataList.Item2.Length != 0)
            {
                statusTextLabel.Text = $"Помилка з БД: {tableDataList.Item2}";
                statusTextLabel.Foreground = Brushes.Red;
                return;
            }

            tableData.ItemsSource = tableDataList.Item1;
            statusTextLabel.Text = tableDataList.Item1.Count!=0 ? $"Останній запит до БД виконано за {sw.ElapsedMilliseconds} мсек. Кількість записів: {tableDataList.Item1.Count}." : "Немає записів в БД за вказаним фільтром";
        }

        private void importButton_Click(object sender, RoutedEventArgs e)
        {
            var importForm = new ImportForm();
            importForm.Show();
        }

    }
}
