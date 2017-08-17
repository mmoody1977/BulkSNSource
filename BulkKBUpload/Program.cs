using BulkKBUpload.AppServices;
using BulkKBUpload.WebServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkKBUpload
{
    class Program
    {
        static void Main(string[] args)
        {
            // Setup Logger
            StreamWriter sw = new StreamWriter("BulkKBToSN.log", true);
            Logger logger = new Logger(sw);
            logger.Log("Script started", logType.Info).Wait();
            /* Used for testing */
            string folderPath = @"\\localhost\C$\kb_upload";
            // Handle Args
            if (args.Length == 1)
            {
                string arg = args[0];
                char[] argChars = arg.ToCharArray();
                if (argChars.Last() == '\\')
                {
                    // Strip off the last character
                    Console.WriteLine("The folder path should not end with a backslash \\. Remove the backslash and try again.");
                    return;
                }

                if (args[0].ToLower().Contains(@":\"))
                {
                    string drive = args[0].Split(':')[0];
                    folderPath = @"\\localhost\" + drive + "$" + args[0].Split(':')[1];
                }
                else
                {
                    folderPath = args[0];
                }
                if (args[0] == "/?" || args[0] == "-?")
                {
                    Console.WriteLine("Please provide a folder path which contains the documents you want to upload to the ServiceNow Knowledge Base. Example: BulkKBUpload \"c:\\docs\"");
                    return;
                }
            }

            KBUpload kbUpload = new KBUpload(logger);
            kbUpload.BulkUpload(folderPath).Wait();
            Console.WriteLine("##########################################################################");
            Console.WriteLine("##########################################################################");
            Console.WriteLine("Bulk upload complete");
            sw.Close();
            sw.Dispose();
            //SN_Service sn_service = new SN_Service("https://instance.service-now.com", "", "!");

            //Task<List<string>> urls = sn_service.uploadImageToSN(@"\\localhost\C$\Users\mmoody\Pictures\Eva\V__4884.jpg", "a558c6d287843900318a276709434df7");
            //foreach (string url in urls.Result)
            //{
            //    Console.WriteLine(url);
            //}
            Console.ReadLine();
        }
    }
}
