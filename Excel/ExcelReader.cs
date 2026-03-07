using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using macros.Models;
using macros.Utils;

namespace macros.Excel
{
    public class ExcelReader
    {
        public List<StudentRow> ReadStudents(string filePath)
        {
            var students = new List<StudentRow>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);

                var headerMap = GetHeaderMap(worksheet);

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    // Пропускаем полностью пустые строки
                    if (IsRowEmpty(row))
                        continue;

                    var student = new StudentRow
                    {
                        RowNumber = row.RowNumber(),
                        LastName = StringCleaner.Clean(row.Cell(headerMap.LastNameColumn).GetString()),
                        FirstName = StringCleaner.Clean(row.Cell(headerMap.FirstNameColumn).GetString()),
                        MiddleName = StringCleaner.Clean(row.Cell(headerMap.MiddleNameColumn).GetString()),
                        Login = StringCleaner.Clean(row.Cell(headerMap.LoginColumn).GetString())
                    };

                    students.Add(student);
                }
            }

            return students;
        }

        private HeaderMap GetHeaderMap(IXLWorksheet worksheet)
        {
            var headerRow = worksheet.Row(1);

            var map = new HeaderMap
            {
                LastNameColumn = FindColumn(headerRow, "Фамилия"),
                FirstNameColumn = FindColumn(headerRow, "Имя"),
                MiddleNameColumn = FindColumn(headerRow, "Отчество"),
                LoginColumn = FindColumn(headerRow, "Логин")
            };

            return map;
        }

        private int FindColumn(IXLRow headerRow, string columnName)
        {
            foreach (var cell in headerRow.CellsUsed())
            {
                var value = StringCleaner.Clean(cell.GetString());

                if (string.Equals(value, columnName, StringComparison.OrdinalIgnoreCase))
                    return cell.Address.ColumnNumber;
            }

            throw new Exception("Не найден обязательный столбец: " + columnName);
        }

        private bool IsRowEmpty(IXLRow row)
        {
            return row.CellsUsed().Count() == 0;
        }

        private class HeaderMap
        {
            public int LastNameColumn { get; set; }
            public int FirstNameColumn { get; set; }
            public int MiddleNameColumn { get; set; }
            public int LoginColumn { get; set; }
        }
    }
}