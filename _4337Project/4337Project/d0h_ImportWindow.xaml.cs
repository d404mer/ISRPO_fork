using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using ExcelDataReader;
using System.Data;



namespace _4337Project
{
    /// <summary>
    /// Interaction logic for d0h_ImportWindow.xaml
    /// </summary>
    public partial class d0h_ImportWindow : Window
    {
        private string connectionString = @"Data Source=DESKTOP-AF0FDGA;Initial Catalog=ISRPO_db;Integrated Security=True;";
        public d0h_ImportWindow()
        {
            InitializeComponent();
            AllowDrop = true;
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string excelFilePath = files[0]; // Берём первый файл
                    ImportExcel(excelFilePath);
                }
            }
        }


        private void ImportExcel(string filePath)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка подключения: {ex.Message}");
                }
            }
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet();
                        DataTable dataTable = result.Tables[0]; // Первый лист Excel

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                            {
                                bulkCopy.DestinationTableName = "Clients";
                                bulkCopy.ColumnMappings.Add(0, "ФИО");
                                bulkCopy.ColumnMappings.Add(1, "Код клиента");
                                bulkCopy.ColumnMappings.Add(2, "Дата рождения");
                                bulkCopy.ColumnMappings.Add(3, "Индекс");
                                bulkCopy.ColumnMappings.Add(4, "Город");
                                bulkCopy.ColumnMappings.Add(5, "Улица");
                                bulkCopy.ColumnMappings.Add(6, "Дом");
                                bulkCopy.ColumnMappings.Add(7, "Квартира");
                                bulkCopy.ColumnMappings.Add(8, "E-mail");

                                bulkCopy.WriteToServer(dataTable);
                            }
                        }
                    }
                }

                MessageBox.Show("Импорт завершен!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка импорта: {ex.Message}");
            }
        }

    }
}
