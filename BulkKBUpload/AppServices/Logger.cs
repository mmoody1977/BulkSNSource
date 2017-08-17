using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkKBUpload.AppServices
{
    public class Logger
    {
        public StreamWriter file { get; set; }
        public Logger(StreamWriter sw)
        {
            file = sw;
        }
        public async Task Log(string message, logType type)
        {
            // bool logWritten = false;
            //using (StreamWriter file = new StreamWriter("SumoToSN.log", true))
            //{
            if (type == logType.Info)
            {
                file.WriteLineAsync(DateTime.Now.ToString() + " - " + message).Wait();

            }
            if (type == logType.Error)
            {
                file.WriteLineAsync(DateTime.Now.ToString() + " [ERROR] - " + message).Wait();
            }
            //logWritten = true;
            // }
        }
    }
    public enum logType
    {
        Info,
        Error
    }
}
