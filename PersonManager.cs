using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;

public class PersonManager
{
    private const string FilePath = "person.json";

    public class Person
    {
        public string Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string MiddleName { get; set; }
        public string Login { get; set; }
    }

    public static void ProcessPersonData(string[][] personArray)
    {
        // Создаём файл person.json, если его нет
        if (!File.Exists(FilePath))
        {
            File.WriteAllText(FilePath, "[]");
            Console.WriteLine("Файл person.json создан.");
        }

        // Читаем существующие данные из JSON
        string jsonContent = File.ReadAllText(FilePath);
        List<Person> existingPersons = JsonSerializer.Deserialize<List<Person>>(jsonContent) ?? new List<Person>();

        // Создаём HashSet для быстрой проверки существующих логинов
        HashSet<string> existingLogins = new HashSet<string>(
            existingPersons.Select(p => p.Login),
            StringComparer.OrdinalIgnoreCase
        );

        // Regex для проверки логина (только латиница и цифры)
        Regex loginRegex = new Regex(@"^[a-zA-Z0-9]+$");

        List<Person> personsToAdd = new List<Person>();

        for (int i = 0; i < personArray.Length; i++)
        {
            string[] row = personArray[i];
            int rowNumber = i + 1; // Номер строки для вывода (начинаем с 1)

            // Проверка: если массив не содержит 5 элементов
            if (row == null || row.Length != 5)
            {
                Console.WriteLine($"Строка {rowNumber} удалена: некорректное количество полей.");
                continue;
            }

            // Удаление лишних пробелов из всех ячеек
            for (int j = 0; j < row.Length; j++)
            {
                if (row[j] != null)
                {
                    // Удаляем пробелы в начале, в конце и множественные пробелы в середине
                    row[j] = Regex.Replace(row[j].Trim(), @"\s+", " ");
                }
            }

            string id = row[0];
            string firstName = row[1];
            string lastName = row[2];
            string middleName = row[3];
            string login = row[4];

            // Проверка: если есть пустое поле
            if (string.IsNullOrWhiteSpace(id) ||
                string.IsNullOrWhiteSpace(firstName) ||
                string.IsNullOrWhiteSpace(lastName) ||
                string.IsNullOrWhiteSpace(middleName) ||
                string.IsNullOrWhiteSpace(login))
            {
                Console.WriteLine($"Строка {rowNumber} удалена: содержит пустое поле.");
                continue;
            }

            // Проверка: логин содержит только латиницу и цифры
            if (!loginRegex.IsMatch(login))
            {
                Console.WriteLine($"Строка {rowNumber} удалена: логин содержит недопустимые символы.");
                continue;
            }

            // Проверка: если логин уже существует в JSON
            if (existingLogins.Contains(login))
            {
                Console.WriteLine($"Строка {rowNumber} удалена: пользователь с логином '{login}' уже существует.");
                continue;
            }

            // Если все проверки пройдены, добавляем нового пользователя
            Person newPerson = new Person
            {
                Id = id,
                FirstName = firstName,
                LastName = lastName,
                MiddleName = middleName,
                Login = login
            };

            personsToAdd.Add(newPerson);
            existingLogins.Add(login); // Добавляем в HashSet для избежания дубликатов в текущей пачке
            Console.WriteLine($"Строка {rowNumber}: пользователь '{firstName} {lastName}' (логин: {login}) добавлен.");
        }

        // Добавляем новых пользователей к существующим
        existingPersons.AddRange(personsToAdd);

        // Записываем обновлённые данные обратно в JSON
        string updatedJson = JsonSerializer.Serialize(existingPersons, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
        File.WriteAllText(FilePath, updatedJson);

        Console.WriteLine($"\nВсего добавлено записей: {personsToAdd.Count}");
        Console.WriteLine($"Общее количество записей в файле: {existingPersons.Count}");
    }
}