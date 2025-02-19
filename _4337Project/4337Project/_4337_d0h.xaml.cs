using System;
using System.Data.SqlClient;
using System.Windows;



namespace _4337Project
{
    /// <summary>
    /// Interaction logic for _4337_d0h.xaml
    /// </summary>
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
     }
 }
