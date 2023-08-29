using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EML2PST
{
    class messages
    {

        public void msgWHITE(string printString)
        {
            writeLog(printString);
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(printString);
            Console.ResetColor();
        }

        public void msgCOMMON(string printString)
        {
            writeLog(printString);
            Console.WriteLine(printString);
        }
        public void msgGREEN(string printString)
        {
            writeLog(printString);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(printString);
            Console.ResetColor();
        }
        public void msgRED(string printString)
        {
            writeLog(printString);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(printString);
            Console.ResetColor();
        }
        public void writeLog(string printString)
        {
            String logsFolder = "logs\\";
            DirectoryInfo logsDirectory = new DirectoryInfo(logsFolder);
            if (!logsDirectory.Exists)
            {
                logsDirectory.Create();
            }
            System.IO.File.AppendAllText(logsFolder + "log" + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Day + ".txt", DateTime.Now.ToString() + " : " + printString + Environment.NewLine);
        }
    }
}
