using System;
using macros.Excel;

namespace macros
{
    public class StudentProcessor
    {
        private readonly ExcelReader _excelReader;
        private readonly PersonManager _personManager;

        public StudentProcessor()
        {
            _excelReader = new ExcelReader();
            _personManager = new PersonManager();
        }

        public (string groupName, string[][] validData) ProcessStudentsFromExcel(string excelFilePath)
        {
            try
            {
                Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
                Console.WriteLine("║     ОБРАБОТКА ДАННЫХ СТУДЕНТОВ ИЗ EXCEL - T-FLEX DOCs        ║");
                Console.WriteLine("╚════════════════════════════════════════════════════════════════╝\n");

                Console.WriteLine($"📂 Чтение файла: {excelFilePath}");
                var (groupName, studentsTable) = _excelReader.ReadStudents(excelFilePath);
                Console.WriteLine($"✓ Файл прочитан успешно");
                Console.WriteLine($"✓ Группа: {groupName}");
                Console.WriteLine($"✓ Найдено записей: {studentsTable.GetLength(0)}\n");

                string[][] personArray = ConvertTableToArray(studentsTable);

                Console.WriteLine("🔄 Начало обработки и валидации данных...\n");
                var (resultGroupName, validData) = _personManager.ProcessPersonData(groupName, personArray);

                Console.WriteLine("\n╔════════════════════════════════════════════════════════════════╗");
                Console.WriteLine("║                    ОБРАБОТКА ЗАВЕРШЕНА                         ║");
                Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");

                return (resultGroupName, validData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ ОШИБКА ПРИ ОБРАБОТКЕ: {ex.Message}");
                Console.WriteLine($"Тип: {ex.GetType().Name}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Внутренняя ошибка: {ex.InnerException.Message}");
                }
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