using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddin.Services
{
    public static class DataService
    {
        /// <summary>
        /// Читает файл по указанному пути и возвращает перечисление строк файла
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        /// <returns>Перечисление всех строк</returns>
        public static IEnumerable<string> GetLines(string path)
        {
            using (StreamReader dataReader = new StreamReader(path))
            {
                while (!dataReader.EndOfStream)
                {
                    var line = dataReader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    yield return line;
                }
            }
        }
    }
}
