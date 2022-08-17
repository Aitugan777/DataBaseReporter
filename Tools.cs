using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Excel = Microsoft.Office.Interop.Excel;

namespace WriteDataBase
{
    public static class Tools
    {
        public static void WriteResult(this Report report)
        {

            // создаём документ
            DocX document = DocX.Create(Program.Configuration.PathToReport);

            List<string> list = new List<string>()
            {
                $"Количество мужчин: {report.MaleGender}",
                $"Количество женщин: {report.FemaleGender}",
                $"Количество мужчин в возрасте 30-40 лет: {report.Male30To40YearsOld}",
                $"Количество стандартных аккаунтов: {report.StandartAccounds}",
                $"Количество премиум аккаунтов: {report.PremiumAccounds}",
                $"Количество женщин с премиум-аккаунтом в возрасте до 30 лет: {report.FemalePremiumUpTo30YearsOld}"
            };

            foreach(string str in list)
            {
                // вставляем параграф и передаём текст
                document.InsertParagraph(str).
                         // устанавливаем шрифт
                         Font("Calibri").
                         // устанавливаем размер шрифта
                         FontSize(14).
                         // устанавливаем цвет
                         Color(Color.Black).
                         // делаем текст жирным
                         Bold().
                         // выравниваем текст по центру
                         Alignment = Alignment.left;
            }

            // сохраняем документ
            document.Save();
        }

        public static bool ConvertExcelToXMLFile()
        {
            try
            {

                string path = Program.Configuration.PathToExcelFile;
                Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
                List<Person> persons = new List<Person>() { };
                for (int i = 0; i < lastCell.Column; i++) //по всем колонкам
                {
                    for (int j = 1; j < lastCell.Row; j++) // по всем строкам
                    {
                        switch (i)
                        {
                            case 0:
                                persons.Add(new Person() { Name = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString() });
                                break;
                            case 1:
                                persons[j - 1].FirstName = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();
                                break;
                            case 2:
                                persons[j - 1].IsMaleGender = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString() == "м";
                                break;
                            case 3:
                                byte age = 0;
                                if (!byte.TryParse(ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString(), out age))
                                    Console.WriteLine($"У персонажа ({i}) {persons[j].Name} {persons[j].FirstName} невозможно определить возраст!");
                                persons[j - 1].Age = age;
                                break;
                            case 4:
                                persons[j - 1].IsPremium = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString().ToLower() == "премиум";
                                break;
                            default:
                                break;
                        }
                    }
                }
                ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                ObjWorkExcel.Quit(); // выйти из экселя
                GC.Collect();

                DataBase db = new DataBase()
                {
                    Persons = persons
                };

                FileInfo fileInfo = new FileInfo(Program.Configuration.PathToXMLFile);
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(DataBase));
                fileInfo.Delete();
                using (FileStream fs = new FileStream(Program.Configuration.PathToXMLFile, FileMode.OpenOrCreate))
                {
                    xmlSerializer.Serialize(fs, db);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool ReadDataBase()
        {
            string path = Program.Configuration.PathToXMLFile;

            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(DataBase));
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    Program.DB = (DataBase)xmlSerializer.Deserialize(fs);
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public static void ServializeConfig()
        {
            Configuration cfg = new Configuration();
            cfg.PathToReport = @"C:\Users\aitug\OneDrive\Рабочий стол\Development\Programs\WriteDataBaseTest\bin\Debug\Report.docx";
            cfg.PathToXMLFile = @"C:\Users\aitug\OneDrive\Рабочий стол\Development\Programs\WriteDataBaseTest\bin\Debug\DataBase.xml";
            cfg.PathToExcelFile = @"C:\Users\aitug\OneDrive\Рабочий стол\Development\Programs\WriteDataBaseTest\bin\Debug\DataBaseExcel.xlsx";

            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Configuration));
            using (FileStream fs = new FileStream("Configuration.xml", FileMode.OpenOrCreate))
            {
                xmlSerializer.Serialize(fs, cfg);
            }
        }
        public static Configuration GetConfig()
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Configuration));
            using (FileStream fs = new FileStream("Configuration.xml", FileMode.OpenOrCreate))
            {
                Configuration cfg = (Configuration)xmlSerializer.Deserialize(fs);
                return cfg;
            }
        }
    }

    public class DataBase
    {
        [XmlArray("Elements"), XmlArrayItem("Element")]
        public List<Person> Persons = new List<Person>();

        public Report MakeReport()
        {
            Report report = new Report();

            foreach(Person person in Persons)
            {
                if (person.IsMaleGender)
                    report.MaleGender++;
                else
                    report.FemaleGender++;

                if (person.IsMaleGender && person.Age >= 30 && person.Age <= 40)
                    report.Male30To40YearsOld++;

                if (person.IsPremium)
                    report.PremiumAccounds++;
                else
                    report.StandartAccounds++;

                if (!person.IsMaleGender && person.IsPremium && person.Age < 30)
                    report.FemalePremiumUpTo30YearsOld++;
            }
            return report;
        }
    }

    public class Person
    {
        public string Name { get; set; }

        public string FirstName { get; set; }

        public byte Age { get; set; }

        public bool IsMaleGender { get; set; }

        public bool IsPremium { get; set; }
        
    }

    public class Configuration
    {
        [XmlAttribute("PathToXMLFile")]
        public string PathToXMLFile;

        [XmlAttribute("PathToPeport")]
        public string PathToReport;

        [XmlAttribute("PathToExcelFile")]
        public string PathToExcelFile;
    }



    public class Report
    {
        /// <summary>
        /// Количество мужчин
        /// </summary>
        public int MaleGender { get; set; }

        /// <summary>
        /// Количество женщин
        /// </summary>
        public int FemaleGender { get; set; }

        /// <summary>
        /// Количество мужчин в возрасте 30-40 лет
        /// </summary>
        public int Male30To40YearsOld { get; set; }

        /// <summary>
        ///Количество стандартных
        /// </summary>
        public int StandartAccounds { get; set; }

        /// <summary>
        ///Количество премиум-аккаунтов
        /// </summary>
        public int PremiumAccounds { get; set; }

        /// <summary>
        /// Количество женщин с премиум-аккаунтом в возрасте до 30 лет
        /// </summary>
        public int FemalePremiumUpTo30YearsOld { get; set; }
    }
}
