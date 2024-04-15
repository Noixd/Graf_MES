using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;

namespace Graf_MES
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string DB = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = Sreamdb.mdb";
        private OleDbConnection connection = new OleDbConnection(DB);

        public MainWindow()
        {
            InitializeComponent();
        }

        public void OpenDB()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM staf";

            OleDbCommand command = new OleDbCommand(query, connection);

            OleDbDataReader reader = command.ExecuteReader();
        }
    }
}
