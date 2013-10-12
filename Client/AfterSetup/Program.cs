using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AfterSetup
{
    class Program
    {
        static void Main(string[] args)
        {

            //HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\ CurrentVersion\ProductId
            string key = FingerPrint.Value();

            Console.WriteLine(key);

            var client = new WebClient();
            
            string html = client.DownloadString("http://www.ikatalan.com/install?clientId=" + key);
            Console.WriteLine(html);

        }
    }
}
