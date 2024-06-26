using OfficeOpenXml;
using System;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;

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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("graf_table");

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
                            worksheet.Column(j).Style.Numberformat.Format = "MM/dd/yyyy";
                        }

                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];

                    }
                }

                command = new OleDbCommand("SELECT * FROM staff", connection);

                // Чтение данных из базы данных
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                worksheet = package.Workbook.Worksheets.Add("staff");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        if (j == 4)
                        {
                            worksheet.Column(j).Style.Numberformat.Format = "MM/dd/yyyy";
                        }

                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];

                    }
                }

                command = new OleDbCommand("SELECT * FROM management_staff", connection);

                // Чтение данных из базы данных
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                worksheet = package.Workbook.Worksheets.Add("management_staff");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        if (j == 3)
                        {
                            worksheet.Column(j).Style.Numberformat.Format = "MM/dd/yyyy";
                        }

                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];

                    }
                }

                command = new OleDbCommand("SELECT * FROM crew_1", connection);

                // Чтение данных из базы данных
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                worksheet = package.Workbook.Worksheets.Add("crew_1");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                       worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                command = new OleDbCommand("SELECT * FROM crew_2", connection);

                // Чтение данных из базы данных
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                worksheet = package.Workbook.Worksheets.Add("crew_2");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                command = new OleDbCommand("SELECT * FROM crew_3", connection);

                // Чтение данных из базы данных
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                worksheet = package.Workbook.Worksheets.Add("crew_3");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }
                // Сохранение файла Excel

                command = new OleDbCommand("SELECT * FROM work_positions", connection);

                // Чтение данных из базы данных
                adapter = new OleDbDataAdapter(command);
                dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Создание нового приложения Excel и книги
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                worksheet = package.Workbook.Worksheets.Add("work_positions");

                // Запись данных в Excel
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                connection.Close();

                package.SaveAs("Export_graf_table.xlsx");
                MessageBox.Show("Данные успешно экспортированы", "Экспорт");
                try
                {
                    Process.Start("Export_graf_table.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при запуске приложения: " + ex.Message);
                }
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

            MessageBoxResult result = MessageBox.Show("Вы уверены, что хотите удалить эти данные?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question);
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
                            querry = "DELETE FROM `crew_1` WHERE `Код` =";
                            break;

                        case 1:
                            querry = "DELETE FROM `crew_2` WHERE `Код` =";
                            break;

                        case 2:
                            querry = "DELETE FROM `crew_3` WHERE `Код` =";
                            break;
                    }

                    OleDbCommand command = new OleDbCommand(querry + id_to_delete, connection);

                    connection.Open();

                    if (command.ExecuteNonQuery() == 1) MessageBox.Show("Данные успешно удалены", "Удаление");

                    else MessageBox.Show("Ошибка", "Удаление");

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

            MessageBoxResult result = MessageBox.Show("Вы уверены, что хотите удалить эти данные?", "Удаление", MessageBoxButton.YesNo, MessageBoxImage.Question);
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

                    if (command.ExecuteNonQuery() == 1) MessageBox.Show("Данные успешно удалены", "Удаление");

                    else MessageBox.Show("Ошибка", "Удаление");

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

            // Проверка, что пользователь завершил редактирование ячейки            
            //MessageBox.Show(e.EditAction.ToString());
            if (e.EditAction.ToString() == "Commit")
            {
                    edit_column = e.Column.Header.ToString();
                    var edit_value = ((TextBox)e.EditingElement).Text.ToString();

                if (((System.Data.DataRowView)e.Row.Item).IsNew)
                {
                    switch (comboBox2.SelectedIndex)
                    {
                        case 0:
                            querry = "INSERT INTO staff (" + edit_column + ") VALUES ('" + edit_value + "')";
                            //MessageBox.Show(querry);
                            break;

                        case 1:
                            querry = "INSERT INTO management_staff (" + edit_column + ") VALUES ('" + edit_value + "')";
                            //MessageBox.Show(querry);
                            break;
                    }
                }

                else
                {
                    //MessageBox.Show(dataGrid4.Items.Count.ToString());
                    edit_row = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[0].ToString());
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
                }
                
            }

            OleDbConnection connection = new OleDbConnection(DB);
            OleDbCommand command = new OleDbCommand(querry, connection);

            connection.Open();

            try
            {
                if (command.ExecuteNonQuery() == 1)
                {
                    MessageBox.Show("Данные изменены", "Изменение");
                }
                else
                {
                    MessageBox.Show("Данные не изменены", "Изменение");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            connection.Close();
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
                        MessageBox.Show("Данные изменены", "Изменение");
                    }
                    else
                    {
                        MessageBox.Show("Данные не изменены", "Изменение");
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

            // Проверка, что пользователь завершил редактирование ячейки
            //MessageBox.Show(e.EditAction.ToString());
            if (e.EditAction.ToString() == "Commit")
            {
                edit_column = e.Column.Header.ToString();
                var edit_value = ((TextBox)e.EditingElement).Text.ToString();

                if (((System.Data.DataRowView)e.Row.Item).IsNew)
                {
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            querry = "INSERT INTO crew_1 (" + edit_column + ") VALUES ('" + edit_value + "')";
                            //MessageBox.Show(querry);
                            break;

                        case 1:
                            querry = "INSERT INTO crew_2 (" + edit_column + ") VALUES ('" + edit_value + "')";
                            //MessageBox.Show(querry);
                            break;

                        case 2:
                            querry = "INSERT INTO crew_3 (" + edit_column + ") VALUES ('" + edit_value + "')";
                            //MessageBox.Show(querry);
                            break;
                    }
                }

                else
                {
                    //MessageBox.Show(dataGrid4.Items.Count.ToString());
                    edit_row = int.Parse(((System.Data.DataRowView)e.Row.Item).Row.ItemArray[0].ToString());
                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            querry = "UPDATE crew_1 SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                            //MessageBox.Show(querry);
                            break;

                        case 1:
                            querry = "UPDATE crew_2 SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                            //MessageBox.Show(querry);
                            break;

                        case 2:
                            querry = "UPDATE crew_2 SET " + edit_column + " = '" + edit_value + "' WHERE Код = " + edit_row;
                            //MessageBox.Show(querry);
                            break;
                    }
                }

            }

            OleDbConnection connection = new OleDbConnection(DB);
            OleDbCommand command = new OleDbCommand(querry, connection);

            connection.Open();

            try
            {
                if (command.ExecuteNonQuery() == 1)
                {
                    MessageBox.Show("Данные изменены", "Изменение");
                }
                else
                {
                    MessageBox.Show("Данные не изменены", "Изменение");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            connection.Close();
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
                //MessageBox.Show(querry);

                OleDbConnection connection = new OleDbConnection(DB);
                OleDbCommand command = new OleDbCommand(querry, connection);

                connection.Open();

                try
                {
                    if (command.ExecuteNonQuery() == 1)
                    {
                        MessageBox.Show("Данные изменены", "Изменение");
                    }
                    else
                    {
                        MessageBox.Show("Данные не изменены", "Изменение");
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

