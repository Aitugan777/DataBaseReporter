using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace WriteDataBase
{
    internal class Program
    {

        public static DataBase DB;

        public static Configuration Configuration;

        static void Main(string[] args)
        {

            try
            {
                Configuration = Tools.GetConfig();
            }
            catch
            {
                Console.WriteLine("Конфигурация создана! Укажите пути для файлов!");
                Tools.ServializeConfig();
            }
            Configuration = Tools.GetConfig();


            if (Tools.ConvertExcelToXMLFile())
            {
                Console.WriteLine($"Конвертация exsel в xml успешно прошла!");

                if (Tools.ReadDataBase())
                {
                    Console.WriteLine($"База данных успешно прочитано!");

                    Tools.WriteResult(DB.MakeReport());

                    Console.WriteLine($"Отчет успешно создан!");
                }
                else
                {
                    Console.WriteLine($"Не удалось найти файл!");
                }
            }
            else
            {
                Console.WriteLine($"Не удалось найти .xlsx файл ({Configuration.PathToExcelFile})!\nУкажите путь в Configuration.xml");
            }
            

            Console.ReadKey();
        }
    }
}
