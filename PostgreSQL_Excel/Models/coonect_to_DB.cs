using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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
            // rruga hap pas hapi per gjetjen e pathit ku ndodhet file exe (setupi) i programit. dhe vendosja aty e file text me te dhenat e databases qe do te perdor programi.
            //    string codeBase = Assembly.GetCallingAssembly().CodeBase;
            //    UriBuilder uri = new UriBuilder(codeBase);
            //    string path = Uri.UnescapeDataString(uri.Path);
            //    string path1= Path.GetDirectoryName(path);

            // rruga e shkurter per te kapur path-in dmth vendodhjen e file exe ne te cilin do vendoset file tekst me te dhenate database q do telidhet programi.

            // path2 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).ToString().Remove(0, 6);
            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase).ToString().Remove(0, 6) + "\\" + "data_connection_of_servers.txt";
          
            // string path = Environment.CurrentDirectory + "/" + "data_connection_of_servers.txt";
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
