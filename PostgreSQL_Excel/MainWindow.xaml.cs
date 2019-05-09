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
using Microsoft.Win32;
using System.Data;
using System.IO;
using System.Data.OleDb;
//using _excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Npgsql;
using PostgreSQL_Excel.Models;

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
        private NpgsqlConnection connection;
        private NpgsqlCommand command, comand1;
        private NpgsqlDataReader dataReader;
        private string NOvalueChange, valueChange;
        private bool data_load_from_excel_file = false;
        private bool file_excel_formated_ok = false;
        List<string> Excel_colums = new List<string>();
        private int row_selected;
        private int excelcopy = 0;
       
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
           { "Kodi(S)", "supplier_code" },
            { "Pershkrim", "name"},
            {"NIPT", "tax_id"},
            { "Qyteti", "city_id"},

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
          
        }
        // funksion qe kontrollon nese file excel eshte item . dmth do shtohettek tabela Item ne database.
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
                    command = new NpgsqlCommand("SELECT * from public.item_unit", connection);
                    NpgsqlDataReader dr = command.ExecuteReader();
                    while(dr.Read())
                    {
                        string name = dr.GetString(2);
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
                conn.Open();
                //  DataTable dt = new DataTable();
              System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,null);
                conn.Close();
                cmb_sheets.Items.Clear();
                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
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
                        OleDbDataAdapter mydataadapter = new OleDbDataAdapter("Select * from [" + cmb_sheets.Text + "]", conn);

                        //System.Data.DataTable
                            dt_Excel = new System.Data.DataTable();
                        mydataadapter.Fill(dt_Excel);
                        conn.Close();
                       // cmb_sheets.Items.Clear();
                        dataGridView1.ItemsSource = dt_Excel.DefaultView;
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

        private void Btn_show_db_selected_on_cmb_Click(object sender, RoutedEventArgs e)
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                try
                {

                    connection = new NpgsqlConnection(connstring);
                    command = new NpgsqlCommand("SELECT * from public.item", connection);
                    NpgsqlDataAdapter NpgsqlDA = new NpgsqlDataAdapter();
                    NpgsqlDA.SelectCommand = command;
                    System.Data.DataTable dbdataset = new System.Data.DataTable();
                    NpgsqlDA.Fill(dbdataset);
                    dataGridView1.ItemsSource = dbdataset.DefaultView;


                }
                catch (Exception msg)
                {
                    //MessageBox.Show("You can't connect with database!Please chek data connections saved in the file and try again! " + "Server=127.0.0.1; Port=5432; User Id=postgres; Password=b2b4cc1b2; Database=DataStudent;");
                    MessageBox.Show(msg.Message);

                }
            }
            else
            {
                MessageBox.Show("Connection to dataBase has Failed Because File with data connections not Exist or name of the file has changed!");
            }

        }

        private void Btn_add_Click(object sender, RoutedEventArgs e)
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
                                    int Item_unit_ID = 0; int Item_sales_tax_percentage = 0;
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
                                            else { Item_sales_tax_percentage = 0; }

                                            connection.Close();


                                        }
                                        catch (Exception ex)
                                        {
                                            //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                            //dt_saved_ok = false;
                                            MessageBox.Show(ex.Message);

                                        }
                                        string Query = "insert into public.item (item_code,name,barcode,price,tax_rate_id,item_unit_id) values('" + dr[0] + "','" + dr[1] + "','" + dr[2] + "','" + dr[3] + "','" + Item_sales_tax_percentage + "','" + Item_unit_ID + "');";
                                        try
                                        {
                                            connection = new NpgsqlConnection(connstring);
                                            command = new NpgsqlCommand(Query, connection);
                                            connection.Open();
                                            dataReader = command.ExecuteReader();
                                            connection.Close();
                                            ////  MessageBox.Show("Data saved to the database!");
                                            //while (dataReader.Read())
                                            //{

                                            //}
                                        }
                                        catch (Exception ex)
                                        {
                                            //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                            //dt_saved_ok = false;
                                            MessageBox.Show(ex.Message);
                                        }
                                    }
                                }
                                else { MessageBox.Show("Connection with Data base failed!"); }

                            }
                            else
                            {
                                MessageBox.Show("Ky file nuk mund te shtohet ne Database! Ju keni zgjedhur {0}. Sigurohu qe keni zgjedhur opsionin e duhur ne te cilin doni te shtoni te dhenat e file Excel.", user_select);
                            }

                        }
                        break;
                    case "Klienti(Customer)":
                    {
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
                                    //zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi - e
                                  
                                    string Query = "insert into public.customer (customer_code,name,tax_id,city_id) values('" + dr["Kodi"] + "','" + dr["Pershkrim"] + "','" + dr["NIPT"] + "','" + City_id + "');";
                                    try
                                    {
                                        connection = new NpgsqlConnection(connstring);
                                        command = new NpgsqlCommand(Query, connection);
                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        ////  MessageBox.Show("Data saved to the database!");
                                        //while (dataReader.Read())
                                        //{

                                        //}
                                    }
                                    catch (Exception ex)
                                    {
                                        //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                        //dt_saved_ok = false;
                                        MessageBox.Show(ex.Message);
                                    }
                                }
                            }
                            else { MessageBox.Show("Connection with Data base failed!"); }

                        }
                        else
                        {
                            MessageBox.Show("Ky file nuk mund te shtohet ne Database! Ju keni zgjedhur {0}. Sigurohu qe keni zgjedhur opsionin e duhur ne te cilin doni te shtoni te dhenat e file Excel.", user_select);
                        }

                    }
                        break;
                    case "Furnizuesi(Supplier)":
                    {
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
                                    //zgjedhja ne database tek tabela item_unit e vleres qe i korespondon kesaj njesi - e

                                    string Query = "insert into public.supplier (supplier_code,name,tax_id,city_id) values('" + dr["Kodi(S)"] + "','" + dr["Pershkrim"] + "','" + dr["NIPT"] + "','" + City_id + "');";
                                    try
                                    {
                                        NpgsqlConnection connection = new NpgsqlConnection(connstring);
                                        NpgsqlCommand command = new NpgsqlCommand(Query, connection);
                                        connection.Open();
                                        dataReader = command.ExecuteReader();
                                        connection.Close();
                                        ////  MessageBox.Show("Data saved to the database!");
                                        //while (dataReader.Read())
                                        //{

                                        //}
                                    }
                                    catch (Exception ex)
                                    {
                                        //MessageBox.Show("You can't connect with database and for this reason you can not save this data!Please chek data connections saved in the file and try again");
                                        //dt_saved_ok = false;
                                        MessageBox.Show(ex.Message);
                                    }
                                }
                            }
                            else { MessageBox.Show("Connection with Data base failed!"); }

                        }
                        else
                        {
                            MessageBox.Show("Ky file nuk mund te shtohet ne Database! Ju keni zgjedhur {0}. Sigurohu qe keni zgjedhur opsionin e duhur ne te cilin doni te shtoni te dhenat e file Excel.", user_select);
                        }

                    }
                        break;
                    default:
                        {
                        MessageBox.Show("File i Gabuar! Sigurohu qe file i zgjedhur eshte ne formatin e duhur!");
                        }
                        break;
                }
               

        }



       
    }
}
