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
    }
}
