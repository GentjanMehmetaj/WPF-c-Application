using System;
using System.Collections.Generic;

using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Data.OleDb;
using Npgsql;
using PostgreSQL_Excel.Models;
using System.Collections.ObjectModel;

namespace PostgreSQL_Excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        List<string> dataItems = new List<string>();
        data_server_connection DtServer = new data_server_connection();
        connect_DB pg_Connect = new connect_DB();
        System.Data.DataTable dt_Excel = new System.Data.DataTable();
        System.Data.DataTable dt_Excel_temp = new System.Data.DataTable();
        private NpgsqlConnection connection;
        private NpgsqlCommand command;
        private NpgsqlDataReader dataReader;
        //variabel qe te bejdallimrasti kur te dhenat jane shfaqur ngad ata base  ose excel file.
        private bool data_load_from_excel_file = false;
        //varibla te deklaruara qe te kapen ne gjithe kete faqe. ne menyre qe nese duam te fshijme apo te bejme update ne 
        // data base te ruajme vleren e qelizes ne momentin e selectimit.
        private int njesia_selectuar = -1;
        int tvsh_selectuar = -1;
        int id_of_selected_row;
        int qytet_selectuar = -1;
        //varialblaglobal qe do perdoren per pagination te te dhenave te nxjerra nga databasa
        DataTable dt_Products = new DataTable();
        private int paging_PageIndex = 1;
        private int paging_NoOfRecPerPage = 28;
        private enum PagingMode { First = 1, Next = 2, Previous = 3, Last = 4,Go = 5 };




        Dictionary<string, string> Item_dictionary = new Dictionary<string, string>()
        {
            { "kodi", "item_code" },
            { "artikulli", "name"},
            {"barkodi", "barcode"},
            { "cmimi", "price"},
            {"tvsh", "tax_rate_id"},
            {"njesi", "item_unit_id"}
           // {"cmimi i blerjes", "buying_price"}

        };
        Dictionary<string, string> customer_dictionary = new Dictionary<string, string>()
        {
            { "Kodi", "customer_code" },
            { "Pershkrim", "name"},
            {"NIPT", "tax_id"},
            { "Qyteti", "city_id"},


        };
        Dictionary<string, string> supplier_dictionary = new Dictionary<string, string>()
        {
           { "Kodi(Furnitor)", "supplier_code" },
            { "Pershkrim(Furnitor)", "name"},
            {"NIPT(Furnitor)", "tax_id"},
            { "Qyteti(Furnitor)", "city_id"},

        };
        Dictionary<string, string> user_select_dictionary = new Dictionary<string, string>()
        {
            { "Artikulli(Item)", "Item_dictionary" },
            { "Klienti(Customer)", "customer_dictionary"},
            {"Furnizuesi(Supplier)", "supplier_dictionary"}
        };
           

        public MainWindow()
        {
            InitializeComponent();
            fill_cmb_njesi();
            fill_cmb_tvsh();

        }
        //selectimi i nje celli ne data grid
        //public static DataGridCell GetCell(DataGrid dataGrid, DataGridRow rowContainer, int column)
        //{
        //    if (rowContainer != null)
        //    {
        //        DataGridCellsPresenter presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
        //        if (presenter == null)
        //        {
        //            /* if the row has been virtualized away, call its ApplyTemplate() method
        //             * to build its visual tree in order for the DataGridCellsPresenter
        //             * and the DataGridCells to be created */
        //            rowContainer.ApplyTemplate();
        //            presenter = FindVisualChild<DataGridCellsPresenter>(rowContainer);
        //        }
        //        if (presenter != null)
        //        {
        //            DataGridCell cell = presenter.ItemContainerGenerator.ContainerFromIndex(column) as DataGridCell;
        //            if (cell == null)
        //            {
        //                /* bring the column into view
        //                 * in case it has been virtualized away */
        //                dataGrid.ScrollIntoView(rowContainer, dataGrid.Columns[column]);
        //                cell = presenter.ItemContainerGenerator.ContainerFromIndex(column) as DataGridCell;
        //            }
        //            return cell;
        //        }
        //    }
        //    return null;
        //}
        void SelectRowByIndex(DataGrid dataGrid, int rowIndex)
        {

            //if (!data_load_from_excel_file)
            //{
            //if (!dataGrid.SelectionUnit.Equals(DataGridSelectionUnit.FullRow))
            //    throw new ArgumentException("The SelectionUnit of the DataGrid must be set to FullRow.");

            //if (rowIndex < 0 || rowIndex > (dataGrid.Items.Count - 1))
            //    throw new ArgumentException(string.Format("{0} is an invalid row index.", rowIndex));
            dataGrid.UnselectAll();
               // dataGrid.SelectedItems.Clear();
                /* set the SelectedItem property */
                object item = dataGrid.Items[rowIndex]; // = Product X
                dataGrid.SelectedItem = item;
          
                //DataGridRow row = dataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
                //if (row == null)
                //{
                //    /* bring the data item (Product object) into view
                //     * in case it has been virtualized away */
                //    dataGrid.ScrollIntoView(item);
                //    row = dataGrid.ItemContainerGenerator.ContainerFromIndex(rowIndex) as DataGridRow;
                //}
                //if (row != null)
                //{
                //    DataGridCell cell = GetCell(dataGrid, row, 0);
                //    if (cell != null)
                //        cell.Focus();
                //}
            //}
        }

        // funksion qe kontrollon nese file excel eshte item . dmth do shtohet tek tabela Item ne database.
        bool excel_file_loaded_is_item(List<string> ex_colums)
        {
            bool check_file = false;
            int i = 0;
            int nr_el_liste = 0;
            if (ex_colums.Count == Item_dictionary.Count)
            {
                nr_el_liste = ex_colums.Count;
                foreach(string el in ex_colums)
                {
                   if(Item_dictionary.ContainsKey(el))
                    {
                        i += 1;
                    }
                }
                if (i == nr_el_liste)
                {
                    check_file = true;
                }
                else
                {
                    check_file = false;
                }
            }
            else {
                check_file = false;
            }
            return check_file;
           
        }
        bool excel_file_loaded_is_customer(List<string> ex_colums)
        {
            bool check_file = false;
            int i = 0;
            int nr_el_liste = 0;
            if (ex_colums.Count == customer_dictionary.Count)
            {
                nr_el_liste = ex_colums.Count;
                foreach (string el in ex_colums)
                {
                    if (customer_dictionary.ContainsKey(el))
                    {
                        i += 1;
                    }
                }
                if (i == nr_el_liste)
                {
                    check_file = true;
                }
                else
                {
                    check_file = false;
                }
            }
            else
            {
                check_file = false;
            }
            return check_file;

        }
        bool excel_file_loaded_is_supplier(List<string> ex_colums)
        {
            bool check_file = false;
            int i = 0;
            int nr_el_liste = 0;
            if (ex_colums.Count == supplier_dictionary.Count)
            {
                nr_el_liste = ex_colums.Count;
                foreach (string el in ex_colums)
                {
                    if (supplier_dictionary.ContainsKey(el))
                    {
                        i += 1;
                    }
                }
                if (i == nr_el_liste)
                {
                    check_file = true;
                }
                else
                {
                    check_file = false;
                }
            }
            else
            {
                check_file = false;
            }
            return check_file;

        }
        //funksion per mbushjen nga database te combobox njesi.
        // nese perdoret ne shtimin manual te te dhenave ne database
        void fill_cmb_njesi()
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                try
                {

                    connection = new NpgsqlConnection(connstring);
                    connection.Open();
                    command = new NpgsqlCommand("SELECT name from public.item_unit", connection);
                    NpgsqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        cmb_njesi.Items.Add(name);
                    }
                    connection.Close();

                }
                catch (Exception msg)
                {
                    //MessageBox.Show("You can't connect with database!Please chek data connections saved in the file and try again! " + "Server=127.0.0.1; Port=5432; User Id=postgres; Password=b2b4cc1b2; Database=DataStudent;");
                    MessageBox.Show(msg.Message);

                }
            }

        }
        void fill_cmb_tvsh()
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                try
                {

                    connection = new NpgsqlConnection(connstring);
                    connection.Open();
                    command = new NpgsqlCommand("SELECT name from public.tax_rate", connection);
                    NpgsqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        cmb_tvsh.Items.Add(name);
                    }
                    connection.Close();

                }
                catch (Exception msg)
                {
                    //MessageBox.Show("You can't connect with database!Please chek data connections saved in the file and try again! " + "Server=127.0.0.1; Port=5432; User Id=postgres; Password=b2b4cc1b2; Database=DataStudent;");
                    MessageBox.Show(msg.Message);

                }
            }

        }
        void select_cmb_tvsh(int tvsh)
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                try
                {

                    connection = new NpgsqlConnection(connstring);
                    connection.Open();
                    command = new NpgsqlCommand("SELECT name from public.tax_rate where tax_rate_id="+tvsh+";", connection);
                    NpgsqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        cmb_tvsh.SelectedItem = name;
                    }
                    connection.Close();

                }
                catch (Exception msg)
                {
                    //MessageBox.Show("You can't connect with database!Please chek data connections saved in the file and try again! " + "Server=127.0.0.1; Port=5432; User Id=postgres; Password=b2b4cc1b2; Database=DataStudent;");
                    MessageBox.Show(msg.Message);

                }
            }

        }
        void select_cmb_njesi(int njesi)
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                try
                {

                    connection = new NpgsqlConnection(connstring);
                    connection.Open();
                    command = new NpgsqlCommand("SELECT name from public.item_unit where item_unit_id=" + njesi + ";", connection);
                    NpgsqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        cmb_njesi.SelectedItem = name;
                    }
                    connection.Close();

                }
                catch (Exception msg)
                {
                    //MessageBox.Show("You can't connect with database!Please chek data connections saved in the file and try again! " + "Server=127.0.0.1; Port=5432; User Id=postgres; Password=b2b4cc1b2; Database=DataStudent;");
                    MessageBox.Show(msg.Message);

                }
            }

        }
        void select_qyteti(int qytet)
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                try
                {

                    connection = new NpgsqlConnection(connstring);
                    connection.Open();
                    command = new NpgsqlCommand("SELECT name from public.city where city_id=" + qytet + ";", connection);
                    NpgsqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        txt_cmimi.Text = name;
                    }
                    connection.Close();

                }
                catch (Exception msg)
                {
                    //MessageBox.Show("You can't connect with database!Please chek data connections saved in the file and try again! " + "Server=127.0.0.1; Port=5432; User Id=postgres; Password=b2b4cc1b2; Database=DataStudent;");
                    MessageBox.Show(msg.Message);

                }
            }

        }



        private void Btn_choose_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.Filter = "Excel File | *.xlsx; *.xls; *.xlsm;";
            var browsefile = openfiledialog1.ShowDialog();
            if (browsefile == true)
            {
                this.txt_path.Text = openfiledialog1.FileName;
                string pathconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txt_path.Text + ";Extended Properties=\"Excel 12.0; HDR=YES;\" ; ";
                OleDbConnection conn = new OleDbConnection(pathconn);
                System.Data.DataTable dt = new System.Data.DataTable();
                try
                {
                    conn.Open();
                    dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                    //  DataTable dt = new DataTable();
               
                cmb_sheets.Items.Clear();
                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString().Replace("$", "");
                    string sheet1 = excelSheets[i];
                    i++;
                }
               
                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    cmb_sheets.Items.Add(excelSheets[j]);
                    // Query each excel sheet.
                }

               
            }
        }

        private void Btn_load_Click(object sender, RoutedEventArgs e)
        {
            data_load_from_excel_file = true;
            if (txt_path.Text != "" && cmb_sheets.Text != "")
            {
               // dataGridView1.DataContext = null;
                string pathconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txt_path.Text + ";Extended Properties=\"Excel 12.0; HDR=YES;\" ; ";
                if (File.Exists(txt_path.Text.ToString()))
                {
                    
                    try
                    {
                        OleDbConnection conn = new OleDbConnection(pathconn);
                        conn.Open();
                        OleDbDataAdapter mydataadapter = new OleDbDataAdapter("Select * from [" + cmb_sheets.Text + "$]", conn);

                        //System.Data.DataTable
                        dt_Excel = new System.Data.DataTable();
                        mydataadapter.Fill(dt_Excel);
                        conn.Close();
                       // cmb_sheets.Items.Clear();
                        dataGridView1.ItemsSource = dt_Excel.DefaultView;
                        //ruhet nje koje data table te excel te loduar ne datagrid ne menyre qe te manipulohetne rastin
                        // ezgjedhjes se tabelave tjera ku do behet shtimi ne database.
                        //dt_Excel_temp = dt_Excel;
                        //dt_Excel.Clear();
                        //hide column
                        //dataGridView1.Columns[2].Visibility = Visibility.Collapsed;
                        // dataGridView1.DataContext= dt;

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else { MessageBox.Show("File not exist at this location!Check path!"); }

               // data_load_from_excel_file = true;

            }
            else { MessageBox.Show("Please choose the file and select the sheet name!"); }
        }

        private void Btn_clear_exel_path_Click(object sender, RoutedEventArgs e)
        {
            txt_path.Clear();
            cmb_sheets.Items.Clear();

        }

        private void Btn_clear_datagrid_Click(object sender, RoutedEventArgs e)
        {
            dataGridView1.ItemsSource = null;
            dt_Excel.Clear();
        }

        private void Btn_show_db_selected_on_cmb_Click(object sender, RoutedEventArgs e)
        {//pastrohen vlerat ne menyre qe ne rastin kur nxirren te dhena nga data basa te mos ngelin si vlera dhe te krijojne konfuzion per perdoruesin.
            txt_path.Clear();
            cmb_sheets.Items.Clear();
           
            data_load_from_excel_file = false;
            string user_select = ((ComboBoxItem)cmb_tabelat_neDB.SelectedItem).Content as string;
            if (user_select != null)
            {
                switch (user_select)
                {
                    case "Artikulli(Item)":
                        {
                            DtServer = pg_Connect.connect_database();
                            string connstring = DtServer.dt_connection;
                            bool conn_True = DtServer.fileExist;
                            if (conn_True)
                            {// do krijohet nje funksion qe do marre si paramaeter connstring dhe string qe permban te dhenat e kerkimit ne tabelen perkatese tek database
                                string item_query = "SELECT item_id,name,price,is_with_tax,tax_rate_id,buying_price,barcode,item_code,item_unit_id from public.item";
                                dt_Products.Clear();
                                
                                dataGridView1.ItemsSource = null;
                                tbxPageNum.Clear();

                                ListProducts(connstring, item_query);
                            }
                            else
                            {
                                MessageBox.Show("Lidhja me dataBase nuk eshte e mundur sepse FILE me te dhenat e DataBase-s me te cilen eshte i lidhur programi nuk Egziston ose eshte ndryshuar emri i tij!");
                            }
                        }
                        break;
                    case "Klienti(Customer)":
                        {
                            DtServer = pg_Connect.connect_database();
                            string connstring = DtServer.dt_connection;
                            bool conn_True = DtServer.fileExist;
                            if (conn_True)
                            {// do krijohet nje funksion qe do marre si paramaeter connstring dhe string qe permban te dhenat e kerkimit ne tabelen perkatese tek database
                                string item_query = "SELECT customer_id,customer_code,tax_id,city_id,name from public.customer";
                                dt_Products.Clear(); 
                                dataGridView1.ItemsSource = null;
                                tbxPageNum.Clear();
                                ListProducts(connstring, item_query);
                               
                            }
                            else
                            {
                                MessageBox.Show("Lidhja me dataBase nuk eshte e mundur sepse FILE me te dhenat e DataBase-s me te cilen eshte i lidhur programi nuk Egziston ose eshte ndryshuar emri i tij!");
                            }
                        }
                        break;
                    case "Furnizuesi(Supplier)":
                        {
                            DtServer = pg_Connect.connect_database();
                            string connstring = DtServer.dt_connection;
                            bool conn_True = DtServer.fileExist;
                            if (conn_True)
                            {// do krijohet nje funksion qe do marre si paramaeter connstring dhe string qe permban te dhenat e kerkimit ne tabelen perkatese tek database
                                string item_query = "SELECT supplier_id,supplier_code,tax_id,city_id,name from public.supplier";
                                dt_Products.Clear(); 
                                dataGridView1.ItemsSource = null;
                                tbxPageNum.Clear();
                                ListProducts(connstring, item_query);

                            }
                            else
                            {
                                MessageBox.Show("Lidhja me dataBase nuk eshte e mundur sepse FILE me te dhenat e DataBase-s me te cilen eshte i lidhur programi nuk Egziston ose eshte ndryshuar emri i tij!");
                            }
                        }
                        break;
                    default:
                        MessageBox.Show("Duhet te zgjidhni njerin nga opsionet.");
                        break;
                }
            }
            else { MessageBox.Show("Asgje per te shfaqur! Duhet te zgjidhni njerin nga opsionet ne menyre qe te vazhdoni kete veprim!"); }

        }

        private void Btn_add_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridView1.ItemsSource != null)
            {
                if (data_load_from_excel_file)
                {
                    //if (dataGridView1.DataContext != null)
                    //{   // zgjedhja mga useri e njerit nga opsionet ne te cilin kerkon qe te beje shtimin ne data base dhe ruajtja ne nje variabel
                    string user_select = ((ComboBoxItem)cmb_tabelat_neDB.SelectedItem).Content as string;
                    
                    // Gjetja tek dictionary e modelit te excelit sipas selectimit te user-it
                    // sipas te cilit do behet shtimi ne data base(sipas rastit ne switch)
                    //(user_select_dictionary[user_select])
                    switch (user_select)
                    {
                        case "Artikulli(Item)":
                            {
                                List<string> Excel_colums = new List<string>();
                                int saved_to_db = 0;
                                foreach (DataColumn dc in dt_Excel.Columns)
                                {
                                    Excel_colums.Add(dc.ColumnName);
                                }
                                bool is_item = excel_file_loaded_is_item(Excel_colums);
                                if (is_item)
                                {
                                    DtServer = pg_Connect.connect_database();
                                    string connstring = DtServer.dt_connection;
                                    bool conn_True = DtServer.fileExist;

                                    if (conn_True)
                                    {

                                        foreach (DataRow dr in dt_Excel.Rows)
                                        {
                                            int Item_unit_ID = 0; int Item_sales_tax_percentage = -1;
                                            // zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi-e
                                            string Query1 = "SELECT item_unit_id from public.item_unit WHERE name ='" + dr["njesi"] + "';";
                                           
                                                try
                                                {
                                                    connection = new NpgsqlConnection(connstring);
                                                    NpgsqlCommand comand1 = new NpgsqlCommand(Query1, connection);
                                                    connection.Open();
                                                    var query1_result = comand1.ExecuteScalar();
                                                    if (query1_result != null)
                                                    {
                                                        Item_unit_ID = Convert.ToInt16(query1_result);
                                                    }
                                                    else { Item_unit_ID = 0; }

                                                    connection.Close();


                                                }
                                                catch (Exception ex)
                                                {
                                                    //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                    //dt_saved_ok = false;
                                                    MessageBox.Show(ex.Message);

                                                }
                                            
                                            //zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi - e
                                            string Query2 = "SELECT tax_rate_id from public.tax_rate WHERE item_sales_tax_percentage ='" + dr["tvsh"] + "';";
                                            try
                                            {
                                                connection = new NpgsqlConnection(connstring);
                                                NpgsqlCommand comand2 = new NpgsqlCommand(Query2, connection);
                                                connection.Open();
                                                var query2_result = comand2.ExecuteScalar();
                                                if (query2_result != null)
                                                {
                                                    Item_sales_tax_percentage = Convert.ToInt16(query2_result);
                                                }
                                                else { Item_sales_tax_percentage = -1; }

                                                connection.Close();


                                            }
                                            catch (Exception ex)
                                            {
                                                //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                //dt_saved_ok = false;
                                                MessageBox.Show(ex.Message);

                                            }
                                            if (Item_unit_ID != 0)
                                            {
                                                if (Item_sales_tax_percentage != -1)
                                                { int tax=-1;
                                                    if (Item_sales_tax_percentage == 1)
                                                    {
                                                        tax = 0;
                                                    }
                                                    else { tax = 1; }


                                                    string Query = "insert into public.item (item_code,name,barcode,price,is_with_tax,tax_rate_id,item_unit_id) values('" + dr["kodi"] + "','" + dr["artikulli"] + "','" + dr["barkodi"] + "','" + dr["cmimi"] + "','"+tax+"','" + Item_sales_tax_percentage + "','" + Item_unit_ID + "');";
                                                    try
                                                    {
                                                        connection = new NpgsqlConnection(connstring);
                                                        command = new NpgsqlCommand(Query, connection);
                                                        connection.Open();
                                                        dataReader = command.ExecuteReader();
                                                        connection.Close();
                                                        saved_to_db += 1;

                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                        //dt_saved_ok = false;
                                                        saved_to_db = 0;
                                                        MessageBox.Show(ex.Message);

                                                    }
                                                }
                                                else
                                                {

                                                    int index = dt_Excel.Rows.IndexOf(dr);
                                                    SelectRowByIndex(dataGridView1, index);
                                                    index += 1;
                                                    MessageBox.Show("Ne rreshtin " + index + " TVSH eshte gabim.Rregullo te dhenat ne menyre qe te shtohen ne database!");
                                                    break;
                                                }
                                            }
                                            else {
                                               
                                                int index = dt_Excel.Rows.IndexOf(dr);
                                                SelectRowByIndex(dataGridView1, index);
                                                index += 1;
                                                MessageBox.Show("Ne rreshtin "+index+ " njesia eshte gabim.Rregullo te dhenat ne menyre qe te shtohen ne database!");
                                                break;
                                            }
                                        }
                                        if (saved_to_db == dt_Excel.Rows.Count)
                                        {
                                            MessageBox.Show("Excel file u shtua me sukses tek tabela Item(Artikulli) ne Database!");
                                            dataGridView1.ItemsSource = null;
                                            dt_Excel.Clear();
                                        }

                                    }
                                    else { MessageBox.Show("Lidhja me DataBase nuk eshte e mundur sepse FILE me te dhenat e DataBase-s me te cilen eshte i lidhur programi nuk Egziston ose eshte ndryshuar emri i tij!"); }

                                }
                                else
                                {
                                    MessageBox.Show("Ky file nuk mund te shtohet ne Database! Ju keni zgjedhur " + user_select + ". Sigurohu qe keni zgjedhur opsionin e duhur ne te cilin doni te shtoni te dhenat e file Excel.");
                                }

                            }
                            break;
                        case "Klienti(Customer)":
                            {
                                List<string> Excel_colums = new List<string>();
                                int saved_to_db = 0;
                                foreach (DataColumn dc in dt_Excel.Columns)
                                {
                                    Excel_colums.Add(dc.ColumnName);
                                }
                                bool is_customer = excel_file_loaded_is_customer(Excel_colums);
                                if (is_customer)
                                {
                                    DtServer = pg_Connect.connect_database();
                                    string connstring = DtServer.dt_connection;
                                    bool conn_True = DtServer.fileExist;

                                    if (conn_True)
                                    {

                                        foreach (DataRow dr in dt_Excel.Rows)
                                        {
                                            int City_id = 0;
                                            // zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi-e
                                            string Query1 = "SELECT city_id from public.city WHERE name ='" + dr["Qyteti"] + "';";
                                            try
                                            {
                                                connection = new NpgsqlConnection(connstring);
                                                NpgsqlCommand comand1 = new NpgsqlCommand(Query1, connection);
                                                connection.Open();
                                                var query1_result = comand1.ExecuteScalar();
                                                if (query1_result != null)
                                                {
                                                    City_id = Convert.ToInt16(query1_result);
                                                }
                                                else { City_id = 0; }

                                                connection.Close();


                                            }
                                            catch (Exception ex)
                                            {
                                                //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                //dt_saved_ok = false;
                                                MessageBox.Show(ex.Message);

                                            }
                                            if (City_id != 0)
                                            {

                                                string Query = "insert into public.customer (customer_code,name,tax_id,city_id) values('" + dr["Kodi"] + "','" + dr["Pershkrim"] + "','" + dr["NIPT"] + "','" + City_id + "');";
                                                try
                                                {
                                                    connection = new NpgsqlConnection(connstring);
                                                    command = new NpgsqlCommand(Query, connection);
                                                    connection.Open();
                                                    dataReader = command.ExecuteReader();
                                                    connection.Close();
                                                    saved_to_db += 1;
                                                }
                                                catch (Exception ex)
                                                {
                                                    //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                    //dt_saved_ok = false;
                                                    saved_to_db = 0;
                                                    MessageBox.Show(ex.Message);
                                                }
                                            }
                                            else
                                            {

                                                int index = dt_Excel.Rows.IndexOf(dr);
                                                SelectRowByIndex(dataGridView1, index);
                                                index += 1;
                                                MessageBox.Show("Ne rreshtin " + index + " Emri i Qyteti-t eshte gabim.Rregullo te dhenat ne menyre qe te shtohen ne database!");
                                                break;
                                            }

                                        }
                                        if (saved_to_db == dt_Excel.Rows.Count)
                                        {
                                            MessageBox.Show("Excel file u shtua me sukses tek tabela customer(Klienti) ne Database!");
                                            dataGridView1.ItemsSource = null;
                                            dt_Excel.Clear();
                                        }
                                    }
                                    else { MessageBox.Show("Lidhja me dataBase nuk eshte e mundur sepse FILE me te dhenat e DataBase-s me te cilen eshte i lidhur programi nuk Egziston ose eshte ndryshuar emri i tij!"); }

                                }
                                else
                                {
                                    MessageBox.Show("Ky file nuk mund te shtohet ne Database! Ju keni zgjedhur " + user_select + ". Sigurohu qe keni zgjedhur opsionin e duhur ne te cilin doni te shtoni te dhenat e file Excel.");
                                }

                            }
                            break;
                        case "Furnizuesi(Supplier)":
                            {
                                List<string> Excel_colums = new List<string>();
                                int saved_to_db = 0;
                                foreach (DataColumn dc in dt_Excel.Columns)
                                {
                                    Excel_colums.Add(dc.ColumnName);
                                }
                                bool is_customer = excel_file_loaded_is_supplier(Excel_colums);
                                if (is_customer)
                                {
                                    DtServer = pg_Connect.connect_database();
                                    string connstring = DtServer.dt_connection;
                                    bool conn_True = DtServer.fileExist;

                                    if (conn_True)
                                    {

                                        foreach (DataRow dr in dt_Excel.Rows)
                                        {
                                            int City_id = 0;
                                            // zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi-e
                                            string Query1 = "SELECT city_id from public.city WHERE name ='" + dr["Qyteti(Furnitor)"] + "';";
                                            try
                                            {
                                                connection = new NpgsqlConnection(connstring);
                                                NpgsqlCommand comand1 = new NpgsqlCommand(Query1, connection);
                                                connection.Open();
                                                var query1_result = comand1.ExecuteScalar();
                                                if (query1_result != null)
                                                {
                                                    City_id = Convert.ToInt16(query1_result);
                                                }
                                                else { City_id = 0; }

                                                connection.Close();


                                            }
                                            catch (Exception ex)
                                            {
                                                //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                //dt_saved_ok = false;
                                                MessageBox.Show(ex.Message);

                                            }
                                            //zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi - e
                                            if (City_id!=0)
                                            {
                                                string Query = "insert into public.supplier (supplier_code,name,tax_id,city_id) values('" + dr["Kodi(Furnitor)"] + "','" + dr["Pershkrim(Furnitor)"] + "','" + dr["NIPT(Furnitor)"] + "','" + City_id + "');";
                                                try
                                                {
                                                    NpgsqlConnection connection = new NpgsqlConnection(connstring);
                                                    NpgsqlCommand command = new NpgsqlCommand(Query, connection);
                                                    connection.Open();
                                                    dataReader = command.ExecuteReader();
                                                    connection.Close();
                                                    saved_to_db += 1;


                                                }
                                                catch (Exception ex)
                                                {
                                                    //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                                    //dt_saved_ok = false;
                                                    saved_to_db = 0;
                                                    MessageBox.Show(ex.Message);
                                                }
                                            }
                                            else
                                            {

                                                int index = dt_Excel.Rows.IndexOf(dr);
                                                SelectRowByIndex(dataGridView1, index);
                                                index += 1;
                                                MessageBox.Show("Ne rreshtin " + index + " Emri i Qyteti-t eshte gabim. Rregullo te dhenat ne menyre qe te shtohen ne database!");
                                                break;
                                            }
                                        }
                                        if (saved_to_db == dt_Excel.Rows.Count)
                                        {
                                            MessageBox.Show("Excel file u shtua me sukses tek tabela supplier(Furnizues) ne Database!");
                                            dataGridView1.ItemsSource = null;
                                            dt_Excel.Clear();
                                        }
                                    }
                                    else { MessageBox.Show("Lidhja me dataBase nuk eshte e mundur sepse FILE me te dhenat e DataBase-s me te cilen eshte i lidhur programi nuk Egziston ose eshte ndryshuar emri i tij!"); }

                                }
                                else
                                {
                                    MessageBox.Show("Ky file nuk mund te shtohet ne Database!Ju keni zgjedhur " + user_select + ". Sigurohu qe keni zgjedhur opsionin e duhur ne te cilin doni te shtoni te dhenat e file Excel");
                                }

                            }
                            break;
                        //default:
                        //    {
                        //        MessageBox.Show("Nuk mund te procedohet me kete zgjedhje. !");
                        //    }
                        //    break;
                    }

                }
                else
                {
                    MessageBox.Show("Keto te dhena nuk mund te shtohen ne Database! Ju duhet te zgjidhni nje file Excel ne menyre qe te kryeni kete veprim!");
                }
            }
            else { MessageBox.Show("Asgje per tu shtuar ne Database!"); }
        }

        private void Btn_delete_selected_row_db_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridView1.ItemsSource != null)
            {
                if (!data_load_from_excel_file && id_of_selected_row > 0)
                {
                    string user_select = ((ComboBoxItem)cmb_tabelat_neDB.SelectedItem).Content as string;

                    switch (user_select)
                    {
                        case "Artikulli(Item)":
                            {
                                DtServer = pg_Connect.connect_database();
                                string connstring = DtServer.dt_connection;
                                bool conn_True = DtServer.fileExist;
                                if (conn_True)
                                {
                                    string Query = "delete from public.item where item_id = '" + id_of_selected_row + "';";
                                    try
                                    {
                                        connection = new NpgsqlConnection(connstring);
                                        command = new NpgsqlCommand(Query, connection);
                                        // NpgsqlDataReader dataReader;

                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        MessageBox.Show("Te dhenat u fshine nga tabela Item(Artikulli) ne Database!");
                                        //fshirja e rreshtit ne grid view pas fshirjes se tij te sukseshme ne database
                                        var itemSource = dataGridView1.ItemsSource as DataView;
                                        itemSource.Delete(dataGridView1.SelectedIndex);
                                        dataGridView1.ItemsSource = itemSource;

                                    }
                                    catch (Exception ex)
                                    {
                                        //  MessageBox.Show("You can't connect with database! And for this reason you can not add data to the databas.Please chek data connections saved in the file and try again");
                                        MessageBox.Show(ex.Message);
                                    }


                                }
                            }
                            break;
                        case "Klienti(Customer)":
                            {
                                DtServer = pg_Connect.connect_database();
                                string connstring = DtServer.dt_connection;
                                bool conn_True = DtServer.fileExist;
                                if (conn_True)
                                {
                                    string Query = "delete from public.customer where customer_id = '" + id_of_selected_row + "';";
                                    try
                                    {
                                        connection = new NpgsqlConnection(connstring);
                                        command = new NpgsqlCommand(Query, connection);
                                        // NpgsqlDataReader dataReader;

                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        MessageBox.Show("Te dhenat u fshine nga tabela customer(Klienti) ne Database!");
                                        //fshirja e rreshtit ne grid view pas fshirjes se tij te sukseshme ne database
                                        var itemSource = dataGridView1.ItemsSource as DataView;
                                        itemSource.Delete(dataGridView1.SelectedIndex);
                                        dataGridView1.ItemsSource = itemSource;

                                    }
                                    catch (Exception ex)
                                    {
                                        //  MessageBox.Show("You can't connect with database! And for this reason you can not add data to the databas.Please chek data connections saved in the file and try again");
                                        MessageBox.Show(ex.Message);
                                    }


                                }
                            }
                            break;
                        case "Furnizuesi(Supplier)":
                            {
                                DtServer = pg_Connect.connect_database();
                                string connstring = DtServer.dt_connection;
                                bool conn_True = DtServer.fileExist;
                                if (conn_True)
                                {
                                    string Query = "delete from public.supplier where supplier_id = '" + id_of_selected_row + "';";
                                    try
                                    {
                                        connection = new NpgsqlConnection(connstring);
                                        command = new NpgsqlCommand(Query, connection);
                                        // NpgsqlDataReader dataReader;

                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        MessageBox.Show("Te dhenat u fshine nga tabela supplier(Furnizuesi) ne Database!");
                                        //fshirja e rreshtit ne grid view pas fshirjes se tij te sukseshme ne database
                                        var itemSource = dataGridView1.ItemsSource as DataView;
                                        itemSource.Delete(dataGridView1.SelectedIndex);
                                        dataGridView1.ItemsSource = itemSource;

                                    }
                                    catch (Exception ex)
                                    {
                                        //  MessageBox.Show("You can't connect with database! And for this reason you can not add data to the databas.Please chek data connections saved in the file and try again");
                                        MessageBox.Show(ex.Message);
                                    }


                                }
                            }
                            break;

                    }
                }
            }
        }

        private void DataGridView1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!data_load_from_excel_file)
            {//ne menyre qe mos te ndryshohen nga perdoruesi te  dhenat e nxjerra nga data base dhe te krijojne problemine manupulimin e te 
                //dhenave te saj behet read only. keshtu qe ato mund te ndryshohen vetem me nje nderveprim te sakte nga useri(delete or update etj.)
                dataGridView1.IsReadOnly = true;
                var dg = sender as DataGrid;
                if (dg == null) return;
                var index = dg.SelectedIndex;
                DataRowView row = dg.SelectedItem as DataRowView;
               
                //ne rastin kur fshihet rreshti ne grid view vlera qe merr id_of_selected_row kalon ne exception. vlera duhet te jete me e madhe 0
                //dhe me e vogel se gjatesia e rreshtave te gridview.
                if (index != -1)
                {
                    id_of_selected_row = Convert.ToInt32(row.Row.ItemArray[0].ToString());
                    string user_select = ((ComboBoxItem)cmb_tabelat_neDB.SelectedItem).Content as string;
                    switch (user_select)
                    {
                        case "Artikulli(Item)":
                                txt_kodi.Text = row[Item_dictionary["kodi"]].ToString();
                                txt_artikulli.Text = row[Item_dictionary["artikulli"]].ToString();
                                txt_barkodi.Text = row[Item_dictionary["barkodi"]].ToString();
                                txt_cmimi.Text = row[Item_dictionary["cmimi"]].ToString();

                            //kontrolli nese njesesia eshte bosh. nese permban dicka ajo do jete nje vlere e futur
                            // sipas rregullit dmthsipas tabeles item_unit ne database.normalisht mund tehiqet si kusht. thjeshte per siguri
                            if (row[Item_dictionary["njesi"]].ToString() != "")
                                {
                                njesia_selectuar = Convert.ToInt32(row[Item_dictionary["njesi"]]);
                                    select_cmb_njesi(njesia_selectuar);
                                    // cmb_njesi.SelectedItem= row[Item_dictionary["njesi"]];
                                }
                            //kontrolli nese tvsh eshte bosh. nese permban dicka ajo do jete nje vlere e futur
                            // sipas rregullit dmth sipas tabeles tax_rate ne database.normalisht mund tehiqet si kusht. thjeshte per siguri
                            if (row[Item_dictionary["tvsh"]].ToString() != "")
                                {
                                tvsh_selectuar = Convert.ToInt32(row[Item_dictionary["tvsh"]]);
                                    select_cmb_tvsh(tvsh_selectuar);
                                }
                            break;
                        case "Klienti(Customer)":
                           
                            txt_kodi.Text = row[customer_dictionary["Kodi"]].ToString();
                            txt_artikulli.Text = row[customer_dictionary["Pershkrim"]].ToString();
                            txt_barkodi.Text = row[customer_dictionary["NIPT"]].ToString();
                            //kontrolli nese qyteti eshte bosh. nese permban dicka ajo do jete nje vlere e futur
                            // sipas rregullit dmth sipas tabeles city ne database. normalisht mund tehiqet si kusht. thjeshte per siguri
                            if (row[customer_dictionary["Qyteti"]].ToString() != "")
                            {
                                qytet_selectuar = Convert.ToInt32(row[customer_dictionary["Qyteti"]]);
                                select_qyteti(qytet_selectuar);
                            }

                            break;
                        case "Furnizuesi(Supplier)":
                            txt_kodi.Text = row[supplier_dictionary["Kodi(Furnitor)"]].ToString();
                            txt_artikulli.Text = row[supplier_dictionary["Pershkrim(Furnitor)"]].ToString();
                            txt_barkodi.Text = row[supplier_dictionary["NIPT(Furnitor)"]].ToString();
                            //kontrolli nese qyteti eshte bosh. nese permban dicka ajo do jete nje vlere e futur
                            // sipas rregullit dmth sipas tabeles city ne database. normalisht mund tehiqet si kusht. thjeshte per siguri
                            if (row[supplier_dictionary["Qyteti(Furnitor)"]].ToString() != "")
                            {
                                qytet_selectuar = Convert.ToInt32(row[supplier_dictionary["Qyteti(Furnitor)"]]);
                                select_qyteti(qytet_selectuar);
                            }

                            break;
                        default :
                            MessageBox.Show("Zgjedhje e panjohur");
                            break;

                    }
                    
                }
            }
           

        }
        // eshte pothuajse e njejta procedure qe do ndiqet si ne rastin e butonit add keshtu qe te mos
        //shkruajme kod dy here por te bejme funksione per selectimin e vlerave ne data base ne tabelen perkatese.

        private void Btn_update_selected_row_db_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridView1.ItemsSource != null)
            {
                if (!data_load_from_excel_file && id_of_selected_row > 0)
                {
                    string user_select = ((ComboBoxItem)cmb_tabelat_neDB.SelectedItem).Content as string;

                    switch (user_select)
                    {
                        case "Artikulli(Item)":
                            {  //duhet te kete nje kushte qe tedhenat e shfaqura tek texboxes te nxjerra pas klikimit ne grid view
                                // jane te pakten njera e ndryshme. ne menyre qe te vazhdohet me update.
                                DtServer = pg_Connect.connect_database();
                                string connstring = DtServer.dt_connection;
                                bool conn_True = DtServer.fileExist;
                                if (conn_True)
                                {
                                    int Item_unit_ID = 0; int Item_sales_tax_percentage = -1; int tax = -1;
                                    if (cmb_njesi.SelectedValue != null)
                                    {
                                        string njesi_selected = cmb_njesi.SelectedValue.ToString();
                                        string Query1 = "SELECT item_unit_id from public.item_unit WHERE name ='" + njesi_selected + "';";

                                        try
                                        {
                                            connection = new NpgsqlConnection(connstring);
                                            NpgsqlCommand comand1 = new NpgsqlCommand(Query1, connection);
                                            connection.Open();
                                            var query1_result = comand1.ExecuteScalar();
                                            if (query1_result != null)
                                            {
                                                Item_unit_ID = Convert.ToInt16(query1_result);
                                            }
                                            else { Item_unit_ID = 0; }

                                            connection.Close();


                                        }
                                        catch (Exception ex)
                                        {
                                            //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                            //dt_saved_ok = false;
                                            MessageBox.Show(ex.Message);

                                        }
                                    }
                                    else { MessageBox.Show("Zgjidhni njesine ne menyre qe te vazhdoni me update-in Artikullit");break; }
                                    if (cmb_tvsh.SelectedValue != null)
                                    {
                                        string user_select_tvsh = cmb_tvsh.SelectedValue.ToString();
                                        string Query2 = "SELECT tax_rate_id from public.tax_rate WHERE name ='" + user_select_tvsh + "';";
                                        try
                                        {
                                            connection = new NpgsqlConnection(connstring);
                                            NpgsqlCommand comand2 = new NpgsqlCommand(Query2, connection);
                                            connection.Open();
                                            var query2_result = comand2.ExecuteScalar();
                                            if (query2_result != null)
                                            {
                                                Item_sales_tax_percentage = Convert.ToInt16(query2_result);
                                            }
                                            else { Item_sales_tax_percentage = -1; }

                                            connection.Close();


                                        }
                                        catch (Exception ex)
                                        {
                                            //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                            //dt_saved_ok = false;
                                            MessageBox.Show(ex.Message);

                                        }
                                    }
                                    else { MessageBox.Show("Zgjidhni TVSH-ne ne menyre qe te vazhdoni me update-in Artikullit"); break; }
                                    //ketu vendoset vlera 0 ose 1 e variblit tax e cila do vihet tek tabela item ne data base.

                                    if (Item_sales_tax_percentage == 1)
                                            {
                                                tax = 0;
                                            }
                                            else { tax = 1; }
                                    //kontrollohet vlera e futur tek cmimi nese eshte e sakte dmth ne formatin double dhe pastaj vazhdohet me Update.
                                    decimal cmimi_modified=0;
                                  
                                    if (decimal.TryParse(txt_cmimi.Text.ToString(), out cmimi_modified))
                                    {
                                        string Query = "update public.item SET  name ='" + this.txt_artikulli.Text + "',item_code ='" + this.txt_kodi.Text + "',barcode ='" + this.txt_barkodi.Text + "', price ='" + cmimi_modified + "', is_with_tax ='" + tax + "', tax_rate_id ='" + Item_sales_tax_percentage + "', item_unit_id ='" + Item_unit_ID + "' where item_id = '" + id_of_selected_row + "';";

                                        try
                                        {
                                            connection = new NpgsqlConnection(connstring);
                                            command = new NpgsqlCommand(Query, connection);
                                            // NpgsqlDataReader dataReader;

                                            connection.Open();
                                            dataReader = command.ExecuteReader();
                                            connection.Close();
                                            MessageBox.Show("Te dhenat u u bene Update tek tabela Item(Artikulli) ne Database!");
                                            //fshirja e rreshtit ne grid view pas fshirjes se tij te sukseshme ne database
                                            //var itemSource = dataGridView1.ItemsSource as DataView;
                                            //itemSource.Delete(dataGridView1.SelectedIndex);
                                            //dataGridView1.ItemsSource = itemSource;

                                        }
                                        catch (Exception ex)
                                        {
                                            //  MessageBox.Show("You can't connect with database! And for this reason you can not add data to the databas.Please chek data connections saved in the file and try again");
                                            MessageBox.Show(ex.Message);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Ju duhet te fusni nje vlere te sakte tek cmimi! Nese numri eshte me presje dhjetore perdorni simbolin '.'(pike): P.Sh numri 123.67(njeqind e njezet e tre pike gjashtedhjete e shtate)");
                                    }

                                }
                            }
                            break;
                        case "Klienti(Customer)":
                            {
                               
                                DtServer = pg_Connect.connect_database();
                                string connstring = DtServer.dt_connection;
                                bool conn_True = DtServer.fileExist;
                                if (conn_True)
                                {
                                    string Query = "delete from public.customer where customer_id = '" + id_of_selected_row + "';";
                                    try
                                    {
                                        connection = new NpgsqlConnection(connstring);
                                        command = new NpgsqlCommand(Query, connection);
                                        // NpgsqlDataReader dataReader;

                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        MessageBox.Show("Te dhenat u fshine nga tabela customer(Klienti) ne Database!");
                                        //fshirja e rreshtit ne grid view pas fshirjes se tij te sukseshme ne database
                                        var itemSource = dataGridView1.ItemsSource as DataView;
                                        itemSource.Delete(dataGridView1.SelectedIndex);
                                        dataGridView1.ItemsSource = itemSource;

                                    }
                                    catch (Exception ex)
                                    {
                                        //  MessageBox.Show("You can't connect with database! And for this reason you can not add data to the databas.Please chek data connections saved in the file and try again");
                                        MessageBox.Show(ex.Message);
                                    }


                                }
                            }
                            break;
                        case "Furnizuesi(Supplier)":
                            {
                                DtServer = pg_Connect.connect_database();
                                string connstring = DtServer.dt_connection;
                                bool conn_True = DtServer.fileExist;
                                if (conn_True)
                                {
                                    string Query = "delete from public.supplier where supplier_id = '" + id_of_selected_row + "';";
                                    try
                                    {
                                        connection = new NpgsqlConnection(connstring);
                                        command = new NpgsqlCommand(Query, connection);
                                        // NpgsqlDataReader dataReader;

                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        MessageBox.Show("Te dhenat u fshine nga tabela supplier(Furnizuesi) ne Database!");
                                        //fshirja e rreshtit ne grid view pas fshirjes se tij te sukseshme ne database
                                        var itemSource = dataGridView1.ItemsSource as DataView;
                                        itemSource.Delete(dataGridView1.SelectedIndex);
                                        dataGridView1.ItemsSource = itemSource;

                                    }
                                    catch (Exception ex)
                                    {
                                        //  MessageBox.Show("You can't connect with database! And for this reason you can not add data to the databas.Please chek data connections saved in the file and try again");
                                        MessageBox.Show(ex.Message);
                                    }


                                }
                            }
                            break;

                    }
                }
            }

        }

        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            CustomPaging((int)PagingMode.Go);   
        }

        private void btnFirst_Click(object sender, RoutedEventArgs e)
        {
            CustomPaging((int)PagingMode.First);
        }
        
      
        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            CustomPaging((int)PagingMode.Next);
        }

        private void btnPrev_Click(object sender, RoutedEventArgs e)
        {
            CustomPaging((int)PagingMode.Previous);
        }

        private void btnLast_Click(object sender, RoutedEventArgs e)
        {
            CustomPaging((int)PagingMode.Last);
        }
        //shaqja e faqes ne gridview sipas kerkeses se user-it
        private void CustomPaging(int mode)
        {
            //There is no need for these variables but i created them just for readability
            int totalRecords = dt_Products.Rows.Count;
            int pageSize = paging_NoOfRecPerPage;

            //If total record count is less than  the page size then return.
            if (totalRecords <= pageSize)
            {
                return;
            }

            switch (mode)
            {
                case (int)PagingMode.Next:
                    if (totalRecords > (paging_PageIndex * pageSize))
                    {
                        DataTable tmpTable = new DataTable();
                        tmpTable = dt_Products.Clone();

                        if (totalRecords >= ((paging_PageIndex * pageSize) + pageSize))
                        {
                            for (int i = paging_PageIndex * pageSize;
                                 i < ((paging_PageIndex * pageSize) + pageSize); i++)
                            {
                                tmpTable.ImportRow(dt_Products.Rows[i]);
                            }
                        }
                        else
                        {
                            for (int i = paging_PageIndex * pageSize; i < totalRecords; i++)
                            {
                                tmpTable.ImportRow(dt_Products.Rows[i]);
                            }
                        }

                        paging_PageIndex += 1;

                        dataGridView1.ItemsSource = tmpTable.DefaultView;
                        tmpTable.Dispose();
                    }
                    break;
                case (int)PagingMode.Previous:
                    if (paging_PageIndex > 1)
                    {
                        DataTable tmpTable = new DataTable();
                        tmpTable = dt_Products.Clone();

                        paging_PageIndex -= 1;

                        for (int i = ((paging_PageIndex * pageSize) - pageSize);
                            i < (paging_PageIndex * pageSize); i++)
                        {
                            tmpTable.ImportRow(dt_Products.Rows[i]);
                        }

                        dataGridView1.ItemsSource = tmpTable.DefaultView;
                        tmpTable.Dispose();
                    }
                    break;
                case (int)PagingMode.First:
                    paging_PageIndex = 2;
                    CustomPaging((int)PagingMode.Previous);
                    break;
                case (int)PagingMode.Last:
                    paging_PageIndex = (totalRecords / pageSize);
                    CustomPaging((int)PagingMode.Next);
                    break;
                case (int)PagingMode.Go:
                    int pageGoNum=0;
                    if (tbxPageNum.Text != null && int.TryParse(tbxPageNum.Text, out pageGoNum))
                    {
                        int pageNum = int.Parse(tbxPageNum.Text);
                        //numri i faqeve
                        int nr_of_pages = 0;
                        if (totalRecords % pageSize == 0)
                        {
                            nr_of_pages = totalRecords / pageSize;
                        }
                        else
                        {
                            nr_of_pages = totalRecords / pageSize + 1;
                        }
                        if (pageNum >= 1 && pageNum <= nr_of_pages)
                        {
                            paging_PageIndex = pageNum-1;
                            CustomPaging((int)PagingMode.Next);
                        }
                    }
                        break;
            }

            DisplayPagingInfo();
        }
        //shfaqja e numrit total te faqeve dhe e faqes te cilen kemi te shfaqur
        private void DisplayPagingInfo()
        {
            //There is no need for these variables but i created them just for readability
            int totalRecords = dt_Products.Rows.Count;
            int pageSize = paging_NoOfRecPerPage;
            //numri i faqeve
            int nr_of_pages = 0;
            if (totalRecords % pageSize == 0)
            {
                nr_of_pages = totalRecords / pageSize;
            }
            else
            {
                nr_of_pages = totalRecords / pageSize + 1;
            }

            string pagingInfo = "Numri total i faqeve per tu shfletuar eshte "+ nr_of_pages+ ". Rekordet e shfaqura nga rreshti " + (((paging_PageIndex - 1) * pageSize) + 1) +" deri tek rreshti " + paging_PageIndex*pageSize ;

            if (dt_Products.Rows.Count < (paging_PageIndex * pageSize))
            {
                pagingInfo = "Numri total i faqeve eshte "+ paging_PageIndex + ". Rekordet e shfaqura nga rreshti " + (((paging_PageIndex - 1) * pageSize) + 1) + " deri tek rreshti " + totalRecords;
            }
            lblPagingInfo.Content = pagingInfo;
            lblPageNumber.Content = paging_PageIndex;
            
        }
        //funksion per nxjerrjen e te dhenave nga tabela perkatese ne data base
        private void ListProducts( string db_connection_data,string db_query)
        {
            NpgsqlCommand cmd = new NpgsqlCommand();
            NpgsqlConnection conn = new NpgsqlConnection();
            conn = new NpgsqlConnection(db_connection_data);
            cmd = new NpgsqlCommand(db_query, conn);
            // command = new NpgsqlCommand("SELECT * from public.item", connection);
            NpgsqlDataAdapter NpgsqlDA = new NpgsqlDataAdapter();
            NpgsqlDA.SelectCommand = cmd;
            try
            {

                paging_PageIndex = 1;  //\\ For default
                NpgsqlDA.Fill(dt_Products);

                if (dt_Products.Rows.Count > 0)
                {
                    DataTable tmpTable = new DataTable();

                    //Copying the schema to the temporary table.
                    tmpTable = dt_Products.Clone();

                    //If total record count is greater than page size then
                    //import records from 0 to pagesize (here 20)
                    //Else import reports from 0 to total record count.
                    if (dt_Products.Rows.Count >= paging_NoOfRecPerPage)
                    {
                        for (int i = 0; i < paging_NoOfRecPerPage; i++)
                        {
                            tmpTable.ImportRow(dt_Products.Rows[i]);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < dt_Products.Rows.Count; i++)
                        {
                            tmpTable.ImportRow(dt_Products.Rows[i]);
                        }
                    }

                    //Bind the table to the gridview.
                    dataGridView1.ItemsSource = tmpTable.DefaultView;
                    DisplayPagingInfo();

                    //Dispose the temporary table.
                    tmpTable.Dispose();
                }
                else
                {
                    MessageBox.Show("Nuk ka te dhena per tu shfaqur!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                NpgsqlDA.Dispose();
                cmd.Dispose();
                conn.Dispose();
                
            }
        }

        // kur behet nje perzgjedhje ne combobox:
        //Artikull(item),Klient(Customer),Furnitor(supplier)
        //private bool handle = true;
        //private void ComboBox_DropDownClosed(object sender, EventArgs e)
        //{
        //    if (handle) Handle();
        //    handle = true;
        //}

        //private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    ComboBox cmb = sender as ComboBox;
        //    handle = !cmb.IsDropDownOpen;
        //    Handle();
        //}

        //private void Handle()
        //{
        //    switch (cmb_tabelat_neDB.SelectedItem.ToString().Split(new string[] { ": " }, StringSplitOptions.None).Last())
        //    {
        //        case "Artikulli(Item)":
        //            //Handle for the first combobox 
        //           // dt_Excel = dt_Excel_temp;
        //            break;
        //        case "Klienti(Customer)":
        //            //Handle for the second combobox

        //           // dt_Excel = dt_Excel_temp;
        //            break;
        //        case "Furnizuesi(Supplier)":
        //            //Handle for the third combobox

        //           // dt_Excel = dt_Excel_temp;
        //            break;
        //    }
        //}




    }
}
