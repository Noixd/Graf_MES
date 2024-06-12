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
using System.Data;



namespace Graf_MES
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string DB = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Graf_DB.accdb";
        private OleDbConnection connection = new OleDbConnection(DB);

        public MainWindow()
        {
            InitializeComponent();
            OpenGrafTable();
            OpenWorkPositionsTable();
            
        }

        public void InitComboBox1()
        {
            comboBox1.SelectedIndex = 0;   // по умолчанию будет выбран второй элемент
            OpenCrew1Table();
        }

        public void InitComboBox2()
        {
            comboBox2.SelectedIndex = 0;   // по умолчанию будет выбран второй элемент
            OpenStaffTable();
        }

        public void OpenCrew1Table()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM crew_1";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "crew_1");

            dataGrid1.ItemsSource = dataSet.Tables["crew_1"].DefaultView;

            connection.Close();
        }

        public void OpenCrew2Table()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM crew_2";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "crew_2");

            dataGrid1.ItemsSource = dataSet.Tables["crew_2"].DefaultView;

            connection.Close();
        }

        public void OpenCrew3Table()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM crew_3";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "crew_3");

            dataGrid1.ItemsSource = dataSet.Tables["crew_3"].DefaultView;

            connection.Close();
        }

        public void OpenGrafTable()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM graf_table";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "graf_table");

            dataGrid3.ItemsSource = dataSet.Tables["graf_table"].DefaultView;

            connection.Close();
        }

        public void OpenStaffTable()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM staff";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "staff");

            dataGrid4.ItemsSource = dataSet.Tables["staff"].DefaultView;

            connection.Close();
        }

        public void OpenManagementStaffTable()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM management_staff";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "management_staff");

            dataGrid4.ItemsSource = dataSet.Tables["management_staff"].DefaultView;

            connection.Close();
        }

        public void OpenWorkPositionsTable()
        {
            OleDbConnection connection = new OleDbConnection(DB);

            connection.Open();

            string query = "SELECT * FROM work_positions";


            OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);

            OleDbCommandBuilder commandBuilder = new OleDbCommandBuilder(adapter);
            DataSet dataSet = new DataSet();

            adapter.Fill(dataSet, "work_positions");

            dataGrid2.ItemsSource = dataSet.Tables["work_positions"].DefaultView;

            connection.Close();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {

        }
        private void comboBox1_Initialized(object sender, EventArgs e)
        {
            InitComboBox1();
        }
        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch(comboBox1.SelectedIndex)
            {
                case 0: 
                    OpenCrew1Table();
                    break;
                case 1:
                    OpenCrew2Table();
                    break;
                case 2:
                    OpenCrew3Table();
                    break;
            }
        }

        private void comboBox2_Initialized(object sender, EventArgs e)
        {
            InitComboBox2();
        }

        private void comboBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (comboBox2.SelectedIndex)
            {
                case 0:
                    OpenStaffTable();
                    break;
                case 1:
                    OpenManagementStaffTable();
                    break;
            }
        }

        private void MenuItemRefresh_Click(object sender, RoutedEventArgs e)
        {
            InitializeComponent();
            OpenGrafTable();
            OpenWorkPositionsTable();
            InitComboBox1();
            InitComboBox2();
        }


        private void MenuItemExport_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MenuItemExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }


    }
}