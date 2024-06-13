using OfficeOpenXml;
using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Controls;

namespace Graf_MES
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string DB = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Graf_DB.accdb";
        //private OleDbConnection connection = new OleDbConnection(DB);


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

        public void Refresh_func()
        {
            InitializeComponent();
            OpenGrafTable();
            OpenWorkPositionsTable();
            InitComboBox1();
            InitComboBox2();
        }

        private void comboBox1_Initialized(object sender, EventArgs e)
        {
            InitComboBox1();
        }
        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (comboBox1.SelectedIndex)
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
            Refresh_func();

        }

        private void MenuItemExport_Click(object sender, RoutedEventArgs e)
        {

            using (OleDbConnection connection = new OleDbConnection(DB))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("SELECT * FROM graf_table", connection);

                // Чтение данных из базы данных
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage();
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        if (j == 2)
                        {
                            worksheet.Column(j).Style.Numberformat.Format = "MM/dd/yyyy hh:mm:ss AM/PM";
                        }

                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];

                    }
                }

                // Сохранение файла Excel

                package.SaveAs("Export_graf_table.xlsx");
                MessageBox.Show("Data exported successfully", "Export");
            }
        }

        private void MenuItemExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ContextMenu_delete_row_DG1_Click(object sender, RoutedEventArgs e)
        {
            int id_to_delete;
            TextBlock text_to_delete;
            string querry = null;

            MessageBoxResult result = MessageBox.Show("Are you sure you want to delete this data?", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                if (dataGrid1.SelectedItem != null)
                {
                    id_to_delete = int.Parse(dataGrid1.SelectedIndex.ToString());
                    text_to_delete = dataGrid1.Columns[0].GetCellContent(dataGrid1.Items[id_to_delete]) as TextBlock;
                    id_to_delete = int.Parse(text_to_delete.Text);
                    //MessageBox.Show(id_to_delete.ToString());

                    OleDbConnection connection = new OleDbConnection(DB);

                    switch (comboBox1.SelectedIndex)
                    {

                        case 0:
                            querry = "DELETE FROM `sta` WHERE `Код` =";
                            break;

                        case 1:
                            querry = "DELETE FROM `management_staff` WHERE `Код` =";
                            break;

                        case 2:
                            querry = "DELETE FROM `management_staff` WHERE `Код` =";
                            break;
                    }

                    OleDbCommand command = new OleDbCommand(querry + id_to_delete, connection);

                    connection.Open();

                    if (command.ExecuteNonQuery() == 1) MessageBox.Show("Data deleted successfully", "Delete");

                    else MessageBox.Show("Error", "Delete");

                    connection.Close();

                    dataGrid4.UnselectAll();

                    Refresh_func();
                }
            }
        }

        private void ContextMenu_delete_row_DG4_Click(object sender, RoutedEventArgs e)
        {
            int id_to_delete;
            TextBlock text_to_delete;
            string querry = null;

            MessageBoxResult result = MessageBox.Show("Are you sure you want to delete this data?", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                if (dataGrid4.SelectedItem != null)
                {
                    id_to_delete = int.Parse(dataGrid4.SelectedIndex.ToString());
                    text_to_delete = dataGrid4.Columns[0].GetCellContent(dataGrid4.Items[id_to_delete]) as TextBlock;
                    id_to_delete = int.Parse(text_to_delete.Text);
                    //MessageBox.Show(id_to_delete.ToString());

                    OleDbConnection connection = new OleDbConnection(DB);

                    switch (comboBox2.SelectedIndex)
                    {

                        case 0:
                            querry = "DELETE FROM `staff` WHERE `Код` = ";
                            break;

                        case 1:
                            querry = "DELETE FROM `management_staff` WHERE `Код` = ";
                            break;
                    }

                    OleDbCommand command = new OleDbCommand(querry + id_to_delete, connection);

                    connection.Open();

                    if (command.ExecuteNonQuery() == 1) MessageBox.Show("Data deleted successfully", "Delete");

                    else MessageBox.Show("Error", "Delete");

                    connection.Close();

                    dataGrid4.UnselectAll();

                    Refresh_func();
                }
            }
        }

        private void dataGrid4_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int edit_row;
            string edit_column;
            string querry = null;
            //string last_name;
            //int crew_num;

            //MessageBox.Show(e.EditAction.ToString());
            if (e.EditAction.ToString() == "Commit")
            {

                edit_row = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[0].ToString());
                //last_name = ((System.Data.DataRowView)e.Row.Item).Row.ItemArray[1].ToString();
                //crew_num = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[4].ToString());
                edit_column = e.Column.Header.ToString();
                var edit_value = ((TextBox)e.EditingElement).Text.ToString();



                switch (comboBox2.SelectedIndex)
                {
                    case 0:
                        querry = "UPDATE staff SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                        //MessageBox.Show(querry);
                        break;

                    case 1:
                        querry = "UPDATE management_staff SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                        //MessageBox.Show(querry);
                        break;
                }

                OleDbConnection connection = new OleDbConnection(DB);
                OleDbCommand command = new OleDbCommand(querry, connection);

                connection.Open();

                try
                {
                    if (command.ExecuteNonQuery() == 1)
                    {
                        MessageBox.Show("Data changed", "Editing");
                    }
                    else
                    {
                        MessageBox.Show("Data has not been changed", "Editing");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                connection.Close();
            }
        }

        private void dataGrid3_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int edit_row;
            string edit_column;
            string querry = null;

            //MessageBox.Show(e.EditAction.ToString());
            if (e.EditAction.ToString() == "Commit")
            {

                edit_row = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[0].ToString());
                edit_column = e.Column.Header.ToString();
                var edit_value = ((TextBox)e.EditingElement).Text.ToString();

                querry = "UPDATE graf_table SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                //MessageBox.Show(querry);

                OleDbConnection connection = new OleDbConnection(DB);
                OleDbCommand command = new OleDbCommand(querry, connection);

                connection.Open();

                try
                {
                    if (command.ExecuteNonQuery() == 1)
                    {
                        MessageBox.Show("Data changed", "Editing");
                    }
                    else
                    {
                        MessageBox.Show("Data has not been changed", "Editing");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                connection.Close();
            }
        }

        private void dataGrid1_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int edit_row;
            string edit_column;
            string querry = null;

            //MessageBox.Show(e.EditAction.ToString());
            if (e.EditAction.ToString() == "Commit")
            {

                edit_row = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[0].ToString());
                edit_column = e.Column.Header.ToString();
                var edit_value = ((TextBox)e.EditingElement).Text.ToString();



                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        querry = "UPDATE crew_1 SET " + edit_column + " = " + edit_value + " WHERE Код = " + edit_row;
                        MessageBox.Show(querry);
                        break;

                    case 1:
                        querry = "UPDATE crew_2 SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                        MessageBox.Show(querry);
                        break;

                    case 2:
                        querry = "UPDATE crew_3 SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                        MessageBox.Show(querry);
                        break;
                }

                OleDbConnection connection = new OleDbConnection(DB);
                OleDbCommand command = new OleDbCommand(querry, connection);

                connection.Open();

                try
                {
                    if (command.ExecuteNonQuery() == 1)
                    {
                        MessageBox.Show("Data changed", "Editing");
                    }
                    else
                    {
                        MessageBox.Show("Data has not been changed", "Editing");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                connection.Close();
            }
        }

        private void dataGrid2_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int edit_row;
            string edit_column;
            string querry = null;

            //MessageBox.Show(e.EditAction.ToString());
            if (e.EditAction.ToString() == "Commit")
            {

                edit_row = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[0].ToString());
                edit_column = e.Column.Header.ToString();
                var edit_value = ((TextBox)e.EditingElement).Text.ToString();

                querry = "UPDATE work_positions SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                MessageBox.Show(querry);

                OleDbConnection connection = new OleDbConnection(DB);
                OleDbCommand command = new OleDbCommand(querry, connection);

                connection.Open();

                try
                {
                    if (command.ExecuteNonQuery() == 1)
                    {
                        MessageBox.Show("Data changed", "Editing");
                    }
                    else
                    {
                        MessageBox.Show("Data has not been changed", "Editing");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                connection.Close();
            }
        }
    }
}