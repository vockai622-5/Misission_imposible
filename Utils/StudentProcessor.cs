using System;
using System.Linq;
using macros.Excel;

namespace macros
{
    public class StudentProcessor
    {
        private readonly ExcelReader _excelReader;

        public StudentProcessor()
        {
            _excelReader = new ExcelReader();
        }

        public (string groupName, string[][] validData) ProcessStudentsFromExcel(string excelFilePath)
        {
            try
            {
                // Шаг 1: Читаем данные из Excel
                Console.WriteLine($"📂 Чтение файла: {excelFilePath}");
                var (groupName, studentsTable) = _excelReader.ReadStudents(excelFilePath);
                Console.WriteLine($"✓ Файл прочитан успешно");
                Console.WriteLine($"✓ Группа: {groupName}");
                Console.WriteLine($"✓ Найдено записей: {studentsTable.GetLength(0)}\n");

                // Шаг 2: Конвертируем двумерный массив в массив массивов
                string[][] personArray = ConvertTableToArray(studentsTable);

                // Шаг 3: Обрабатываем данные через PersonManager
                Console.WriteLine("🔄 Начало обработки и валидации данных...\n");
                var personManager = new PersonManager();
                var (resultGroupName, validData) = personManager.ProcessPersonData(groupName, personArray);

                return (resultGroupName, validData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОШИБКА ПРИ ОБРАБОТКЕ: {ex.Message}");
                throw;
            }
        }
        private string[][] ConvertTableToArray(string[,] table)
        {
            int rows = table.GetLength(0);
            int cols = table.GetLength(1);

            string[][] result = new string[rows][];

            for (int i = 0; i < rows; i++)
            {
                result[i] = new string[cols];
                for (int j = 0; j < cols; j++)
                {
                    result[i][j] = table[i, j];
                }
            }

            return result;
        }
    }
}