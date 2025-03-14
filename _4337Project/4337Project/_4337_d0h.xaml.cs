using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;

namespace _4337Project
{
    public partial class _4337_d0h : Window
    {
        private string connectionString = @"Data Source=DESKTOP-AF0FDGA;Initial Catalog=ISRPO_db;Integrated Security=True;";


        public _4337_d0h()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            d0h_ImportWindow import = new d0h_ImportWindow();
            import.Show();
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {


            try
            {
                List<Client> clients = LoadClientsFromDatabase();
                if (clients.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта.");
                    return;
                }

                string filePath = GetSaveFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return;
                }

                SaveToExcel(clients, filePath);
                MessageBox.Show("Данные успешно экспортированы в Excel!");

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}");
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "DELETE FROM Clients";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.ExecuteNonQuery();
                        MessageBox.Show("Таблица успешно очищена!");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}");
                }
            }
        }

        private List<Client> LoadClientsFromDatabase()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Client> clients = new List<Client>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT ClientCode, FIO, Email, Street FROM Clients";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            clients.Add(new Client
                            {
                                ClientCode = reader.GetInt32(0),
                                FIO = reader.GetString(1),
                                Email = reader.GetString(2),
                                Street = reader.GetString(3)
                            });
                        }
                    }
                }
            }

            return clients;
        }

        private void SaveToExcel(List<Client> clients, string filePath)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var groupedByStreet = clients.GroupBy(c => c.Street);

                foreach (var group in groupedByStreet)
                {
                    var sheet = excel.Workbook.Worksheets.Add(group.Key);

                    sheet.Cells[1, 1].Value = "Код клиента";
                    sheet.Cells[1, 2].Value = "ФИО";
                    sheet.Cells[1, 3].Value = "E-mail";

                    var sortedClients = group.OrderBy(c => c.FIO).ToList();
                    for (int i = 0; i < sortedClients.Count; i++)
                    {
                        sheet.Cells[i + 2, 1].Value = sortedClients[i].ClientCode;
                        sheet.Cells[i + 2, 2].Value = sortedClients[i].FIO;
                        sheet.Cells[i + 2, 3].Value = sortedClients[i].Email;
                    }

                    sheet.Cells.AutoFitColumns();
                }

                File.WriteAllBytes(filePath, excel.GetAsByteArray());
            }


        }

        private string GetSaveFilePath()
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                FileName = "ClientsExport",
                DefaultExt = ".xlsx",
                Filter = "Excel files (.xlsx)|*.xlsx"
            };

            bool? result = dlg.ShowDialog();
            return result == true ? dlg.FileName : null;
        }

        private void jsonImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string jsonContent = File.ReadAllText(openFileDialog.FileName);

                    if (string.IsNullOrWhiteSpace(jsonContent))
                    {
                        MessageBox.Show("Выбранный файл пуст!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Преобразование ключей перед десериализацией
                    jsonContent = jsonContent
                        .Replace("CodeClient", "ClientCode")
                        .Replace("FullName", "FIO")
                        .Replace("E_mail", "Email");

                    List<Client> clients = JsonConvert.DeserializeObject<List<Client>>(jsonContent);

                    if (clients == null || clients.Count == 0)
                    {
                        MessageBox.Show("Файл JSON не содержит данных или формат некорректен.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Проверка данных перед добавлением
                    foreach (var client in clients)
                    {
                        if (client.ClientCode <= 0 ||
                            string.IsNullOrWhiteSpace(client.FIO) ||
                            string.IsNullOrWhiteSpace(client.Email) ||
                            string.IsNullOrWhiteSpace(client.Street))
                        {
                            MessageBox.Show("Обнаружены некорректные данные в JSON.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                    }

                    SaveClientsToDatabase(clients);
                    MessageBox.Show("Данные успешно импортированы!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка импорта: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        private void SaveClientsToDatabase(List<Client> clients)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                foreach (var client in clients)
                {
                    string query = "INSERT INTO Clients (ClientCode, FIO, Email, Street) VALUES (@ClientCode, @FIO, @Email, @Street)";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@ClientCode", client.ClientCode);

                        if (string.IsNullOrEmpty(client.FIO))
                            command.Parameters.AddWithValue("@FIO", DBNull.Value);
                        else
                            command.Parameters.AddWithValue("@FIO", client.FIO);

                        if (string.IsNullOrEmpty(client.Email))
                            command.Parameters.AddWithValue("@Email", DBNull.Value);
                        else
                            command.Parameters.AddWithValue("@Email", client.Email);

                        if (string.IsNullOrEmpty(client.Street))
                            command.Parameters.AddWithValue("@Street", DBNull.Value);
                        else
                            command.Parameters.AddWithValue("@Street", client.Street);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }




        private void jsonExport_Click(object sender, RoutedEventArgs e)
        {
            List<Client> clients = LoadClientsFromDatabase();
            var groupedClients = clients.GroupBy(c => c.Street);

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word Documents (*.docx)|*.docx"
            };

            if(saveFileDialog.ShowDialog() == true)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Create(saveFileDialog.FileName, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = document.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = new Body();

                    foreach (var group in groupedClients)
                    {
                        body.Append(new Paragraph(new Run(new Text($"Street: {group.Key}")) { RunProperties = new RunProperties(new Bold()) }));

                        foreach (var client in group)
                        {
                            body.Append(new Paragraph(new Run(new Text($"{client.FIO}, {client.Email}"))));
                        }
                    }

                    mainPart.Document.Append(body);
                    mainPart.Document.Save();
                }

                MessageBox.Show("Данные экспортированы в Word");
            }
        }


        public class Client
        {
            public int ClientCode { get; set; }
            public string FIO { get; set; }
            public string Email { get; set; }
            public string Street { get; set; }
        }
    }
}
