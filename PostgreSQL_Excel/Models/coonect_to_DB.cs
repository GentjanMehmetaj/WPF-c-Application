using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PostgreSQL_Excel.Models
{
    public struct data_server_connection
    {
        public bool fileExist;
        public string dt_connection;
    }


    public class connect_DB
    {

       
        public void Connect_to_DB()
        {
        }

        public data_server_connection connect_database()
        {
            data_server_connection dt_server;
            string path = Environment.CurrentDirectory + "/" + "data_connection_of_servers.txt";
            if (File.Exists(path))
            {
                using (StreamReader str = new StreamReader(path))
                {
                    dt_server.dt_connection = str.ReadLine();
                    dt_server.fileExist = true;
                }

            }
            else
            {
                dt_server.fileExist = false;
                dt_server.dt_connection = null;
            }
            return dt_server;
        }


    }

}
