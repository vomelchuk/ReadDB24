using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media.Imaging;
using Ookii.Dialogs.Wpf;
using ReadDB24.Models;
using ReadDB24.Services;

namespace ReadDB24
{
    /// <summary>
    /// Interaction logic for ImportForm.xaml
    /// </summary>
    public partial class ImportForm : Window
    {
        private SqliteDBService? sqliteDBService;
        private readonly string dataBasePath = $"{Environment.CurrentDirectory}\\24.db";
        private ConfigModel? importConfig;
        private FileIOService fileIOService;

        public ImportForm()
        {
            InitializeComponent();
        }

        private BindingList<MappingExcelModel> ConvertJsonToTable(List<MappingExcelModel> data)
        {
            var result = new BindingList<MappingExcelModel>();
            foreach (var item in data)
            {
                result.Add(item);
            }

            return result;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            fileIOService = new FileIOService($"{Environment.CurrentDirectory}\\config\\mappedExcelFields.json");
            importConfig = fileIOService.LoadMappedFieldsConfig();
            
            var main = ConvertJsonToTable(importConfig.Main);
            var attached = ConvertJsonToTable(importConfig.Attached);
            var reserved = ConvertJsonToTable(importConfig.Reserved);

            sheetMain.ItemsSource = main;
            sheetAttached.ItemsSource = attached;
            sheetReserved.ItemsSource = reserved;
        }

        private void ShowFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var form = new VistaOpenFileDialog();
            var result = form.ShowDialog();
            if (result == true)
            {
                excelFilePath.Text = form.FileName;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (excelFilePath.Text.Length == 0)
            {
                importLog.Text = "Не вибрано Excel файл для імпорту";
                return;
            }

            sqliteDBService = new SqliteDBService(dataBasePath);
            importLog.Text = sqliteDBService.ImportExcel(excelFilePath.Text, importConfig);
            sqliteDBService = null;
            fileIOService.CreateArchive(dataBasePath);
        }
    }
}
