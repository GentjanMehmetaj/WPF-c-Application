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
        private NpgsqlCommand command,comand1;
        private NpgsqlDataReader dataReader;
        private string NOvalueChange, valueChange;
        private bool data_load_from_excel_file = false;
        private bool file_excel_formated_ok = false;

        private int row_selected;
        private int excelcopy = 0;
        private bool Validate_File()
        {

            // ValidFileCheck vlfchek = new ValidFileCheck();
            //  vlfchek.row = 0;
            bool Value = false;
            int row_check = 0;
            if (dataGridView1.Columns.Count == 6)
            {
                //MessageBox.Show(dataGridView1.Columns[0].Name.ToString());
                if (dataGridView1.Columns[0].Header.ToString() == "kodi" && dataGridView1.Columns[1].Header.ToString() == "artikulli" && dataGridView1.Columns[2].Header.ToString() == "barkodi" && dataGridView1.Columns[3].Header.ToString() == "Cmimi me tvsh")
                {
                    file_excel_formated_ok = true;
                    for (int i = 0; i < dataGridView1.Items.Count - 1; i++)
                    {
                        //string str_cell = "";
                        //bool allLetters = str_cell.All(c => Char.IsLetter(c));
                        // bool checknr = dataGridView1.Rows[i].Cells[0].Value.ToString().All(char.IsDigit);
                        //if (dataGridView1.Items[i].Cells[0].Value.ToString != "" && dataGridView1.Rows[i].Cells[0].Value.ToString().All(c => char.IsLetter(c)) && dataGridView1.Rows[i].Cells[1].Value.ToString() != "" && dataGridView1.Rows[i].Cells[1].Value.ToString().All(c => char.IsLetter(c)) && dataGridView1.Rows[i].Cells[2].Value.ToString() != "" && dataGridView1.Rows[i].Cells[2].Value.ToString().All(c => char.IsLetter(c)) && dataGridView1.Rows[i].Cells[3].Value.ToString() != "" && dataGridView1.Rows[i].Cells[3].Value.ToString().All(c => char.IsLetter(c)))
                        //{
                        //    row_check += 1;
                        //    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Aqua;
                        //}
                        //else
                        //{
                        //    // vlfchek.row = row_check + 1;
                        //    dataGridView1.Items[i].DefaultCellStyle.BackColor = Color.Red;
                        //}

                        if (row_check == (dataGridView1.Items.Count - 1))
                        {
                            // vlfchek.ValidFile = true;
                            Value = true;
                        }
                        else
                        {
                            // vlfchek.ValidFile = false;
                            Value = false;

                        }
                    }
                }
                else
                {
                    file_excel_formated_ok = false;
                    MessageBox.Show("File Excel must have colums: kodi,artikulli,barkodi,cmimi me tvsh,cmimi pa tvsh,tvsh,njesi in this range!Yuo have select wrong file(or colums name are wrong");
                }
            }


            return Value;//return true/or false ne varsi te te dhenave ne file te shfaqura ne gridview.
        }
        public MainWindow()
        {
            InitializeComponent();
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

       private void Btn_add_Click(object sender, RoutedEventArgs e)
        {
            DtServer = pg_Connect.connect_database();
            string connstring = DtServer.dt_connection;
            bool conn_True = DtServer.fileExist;
            if (conn_True)
            {
                int Item_unit_ID = 0;
                foreach (DataRow dr in dt_Excel.Rows)
                {
//                    SELECT last_name, first_name
//FROM customer
//WHERE first_name = 'Jamie';
                    
                    string Query1 = "SELECT item_unit_id from public.item_unit WHERE name ='" + dr[6] + "';";                  
                    try
                    {
                        connection = new NpgsqlConnection(connstring);
                        comand1 = new NpgsqlCommand(Query1, connection);
                        connection.Open();
                        var query1_result = comand1.ExecuteScalar();
                        if (query1_result != null)
                        {
                            Item_unit_ID = Convert.ToInt16(query1_result);
                        }
                        else { Item_unit_ID = 0; }
                        
                        connection.Close();
                        
                        //// dataReader["item_unit_id"];

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
                    string Query = "insert into public.item (item_code,name,barcode,price,item_unit_id) values('" + dr[0] + "','" + dr[1] + "','" + dr[2] + "','" + dr[3] + "','" + Item_unit_ID + "');";
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
            else { MessageBox.Show("connection failed!"); }
        }

        private void Btn_show_Click(object sender, RoutedEventArgs e)
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
                    //BindingSource bsource = new BindingSource();
                    //bsource.DataSource = dbdataset;
                    //dataGridView1.DataSource = bsource;
                    //NpgsqlDA.Update(dbdataset);
                    //dataGridView1.AllowUserToAddRows = false;
                    data_load_from_excel_file = false;

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

       
    }
}
