using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using macros.Utils;

namespace macros.Excel
{
    public class ExcelReader
    {
        public (string GroupName, string[,] StudentsTable) ReadStudents(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Не указан путь к файлу.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Файл не найден.", filePath);

            var groupName = StringCleaner.CleanGroupName(
                Path.GetFileNameWithoutExtension(filePath));

            var students = new List<string[]>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var headerMap = GetHeaderMap(worksheet);

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    if (IsRowEmpty(row))
                        continue;

                    // Читаем "№", но не используем в результате
                    var number = GetCellValueSafe(row, headerMap.NumberColumn);

                    var lastName = StringCleaner.CleanName(
                        GetCellValueSafe(row, headerMap.LastNameColumn));

                    var firstName = StringCleaner.CleanName(
                        GetCellValueSafe(row, headerMap.FirstNameColumn));

                    var middleName = StringCleaner.CleanName(
                        GetCellValueSafe(row, headerMap.MiddleNameColumn));

                    var login = StringCleaner.CleanLogin(
                        GetCellValueSafe(row, headerMap.LoginColumn));

                    var shortName = StringCleaner.BuildShortName(
                        lastName,
                        firstName,
                        middleName);

                    if (string.IsNullOrWhiteSpace(lastName)
                        && string.IsNullOrWhiteSpace(firstName)
                        && string.IsNullOrWhiteSpace(middleName)
                        && string.IsNullOrWhiteSpace(login))
                    {
                        continue;
                    }

                    students.Add(new[]
                    {
                        lastName,
                        firstName,
                        middleName,
                        login,
                        shortName
                    });
                }
            }

            return (groupName, BuildStudentsTable(students));
        }

        private HeaderMap GetHeaderMap(IXLWorksheet worksheet)
        {
            var headerRow = worksheet.Row(1);

            return new HeaderMap
            {
                NumberColumn = FindRequiredColumn(headerRow, "№", "N", "No", "Номер"),
                LastNameColumn = FindRequiredColumn(headerRow, "Фамилия"),
                FirstNameColumn = FindRequiredColumn(headerRow, "Имя"),
                MiddleNameColumn = FindRequiredColumn(headerRow, "Отчество"),
                LoginColumn = FindRequiredColumn(headerRow, "Логин")
            };
        }

        private int FindRequiredColumn(IXLRow headerRow, params string[] possibleNames)
        {
            foreach (var cell in headerRow.CellsUsed())
            {
                var value = StringCleaner.Clean(cell.GetString());

                if (possibleNames.Any(name =>
                    string.Equals(value, name, StringComparison.OrdinalIgnoreCase)))
                {
                    return cell.Address.ColumnNumber;
                }
            }

            throw new Exception("Не найден обязательный столбец: " +
                                string.Join(" / ", possibleNames));
        }

        private string GetCellValueSafe(IXLRow row, int columnNumber)
        {
            if (columnNumber <= 0)
                return string.Empty;

            return row.Cell(columnNumber).GetString();
        }

        private bool IsRowEmpty(IXLRow row)
        {
            return row.Cells().All(c => string.IsNullOrWhiteSpace(c.GetString()));
        }

        private string[,] BuildStudentsTable(List<string[]> students)
        {
            const int columnsCount = 5;
            var table = new string[students.Count, columnsCount];

            for (int i = 0; i < students.Count; i++)
            {
                for (int j = 0; j < columnsCount; j++)
                {
                    table[i, j] = students[i][j];
                }
            }

            return table;
        }

        private class HeaderMap
        {
            public int NumberColumn { get; set; }
            public int LastNameColumn { get; set; }
            public int FirstNameColumn { get; set; }
            public int MiddleNameColumn { get; set; }
            public int LoginColumn { get; set; }
        }
    }
}