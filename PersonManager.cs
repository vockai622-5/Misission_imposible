using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace macros
{
    public class PersonManager
    {
        private const string FileName = "person1.json";
        private readonly Random _random;

        public class Person
        {
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string MiddleName { get; set; }
            public string Login { get; set; }
            public string ShortName { get; set; }
            public string Password { get; set; }
            public string Group { get; set; }
        }

        public PersonManager()
        {
            _random = new Random();
        }

        private string GeneratePassword()
        {
            return _random.Next(10000, 99999).ToString();
        }

        private string GetSafeFilePath()
        {
            // Вариант 1: Папка рядом с exe
            try
            {
                string exePath = AppDomain.CurrentDomain.BaseDirectory;
                
                Console.WriteLine($"Попытка 1: Папка с exe - {exePath}");

                if (!exePath.Contains("OneDrive"))
                {
                    string testFile = Path.Combine(exePath, "test_write.tmp");
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);
                    Console.WriteLine("✓ Папка с exe доступна для записи");
                    return Path.Combine(exePath, FileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Не удалось использовать папку с exe: {ex.Message}");
            }

            // Вариант 2: AppData Local
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string projectFolder = Path.Combine(appDataPath, "TFlexDocs");
                Console.WriteLine($"Попытка 2: AppData Local - {projectFolder}");

                if (!Directory.Exists(projectFolder))
                {
                    Directory.CreateDirectory(projectFolder);
                    Console.WriteLine("✓ Папка создана в AppData");
                }

                return Path.Combine(projectFolder, FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Не удалось использовать AppData: {ex.Message}");
            }

            // Вариант 3: Temp
            try
            {
                string tempPath = Path.GetTempPath();
                string projectFolder = Path.Combine(tempPath, "TFlexDocs");
                Console.WriteLine($"П��пытка 3: Temp - {projectFolder}");

                if (!Directory.Exists(projectFolder))
                {
                    Directory.CreateDirectory(projectFolder);
                    Console.WriteLine("✓ Папка создана в Temp");
                }

                return Path.Combine(projectFolder, FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Не удалось использовать Temp: {ex.Message}");
            }

            // Вариант 4: Корень C:
            try
            {
                string rootFolder = @"C:\TFlexDocs";
                Console.WriteLine($"Попытка 4: Корень диска C - {rootFolder}");

                if (!Directory.Exists(rootFolder))
                {
                    Directory.CreateDirectory(rootFolder);
                    Console.WriteLine("✓ Папка создана в корне C:");
                }

                return Path.Combine(rootFolder, FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Не удалось использовать корень C: {ex.Message}");
            }

            throw new Exception("Не удалось найти доступную папку для сохранения файла!");
        }

        public (string groupName, string[][] validData) ProcessPersonData(string groupName, string[][] personArray)
        {
            string filePath = null;

            try
            {
                Console.WriteLine("=== Поиск безопасного места для файла ===\n");
                filePath = GetSafeFilePath();
                Console.WriteLine($"\n✓ Выбран путь: {filePath}\n");

                if (!File.Exists(filePath))
                {
                    File.WriteAllText(filePath, "[]", System.Text.Encoding.UTF8);
                    Console.WriteLine($"✓ Файл {FileName} создан");
                }
                else
                {
                    Console.WriteLine($"✓ Файл {FileName} уже существует");
                }

                string jsonContent = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                List<Person> existingPersons = JsonSerializer.Deserialize<List<Person>>(jsonContent) ?? new List<Person>();

                HashSet<string> existingLogins = new HashSet<string>(
                    existingPersons.Select(p => p.Login),
                    StringComparer.OrdinalIgnoreCase
                );

                Regex loginRegex = new Regex(@"^[a-zA-Z0-9]+$");

                List<Person> personsToAdd = new List<Person>();
                List<string[]> validRows = new List<string[]>();

                Console.WriteLine($"\n=== Обработка группы: {groupName} ===\n");

                for (int i = 0; i < personArray.Length; i++)
                {
                    string[] row = personArray[i];
                    int rowNumber = i + 1;

                    if (row == null || row.Length != 5)
                    {
                        Console.WriteLine($"Строка {rowNumber} удалена: некорректное количество полей.");
                        continue;
                    }

                    for (int j = 0; j < row.Length; j++)
                    {
                        if (row[j] != null)
                        {
                            row[j] = Regex.Replace(row[j].Trim(), @"\s+", " ");
                        }
                    }

                    string lastName = row[0];
                    string firstName = row[1];
                    string middleName = row[2];
                    string login = row[3];
                    string shortName = row[4];

                    if (string.IsNullOrWhiteSpace(lastName) ||
                        string.IsNullOrWhiteSpace(firstName) ||
                        string.IsNullOrWhiteSpace(middleName) ||
                        string.IsNullOrWhiteSpace(login))
                    {
                        Console.WriteLine($"Строка {rowNumber} удалена: содержит пустое поле.");
                        continue;
                    }

                    if (!loginRegex.IsMatch(login))
                    {
                        Console.WriteLine($"Строка {rowNumber} удалена: логин содержит недопустимые символы.");
                        continue;
                    }

                    if (existingLogins.Contains(login))
                    {
                        Console.WriteLine($"Строка {rowNumber} удалена: пользователь с логином '{login}' уже существует.");
                        continue;
                    }

                    string password = GeneratePassword();

                    Person newPerson = new Person
                    {
                        LastName = lastName,
                        FirstName = firstName,
                        MiddleName = middleName,
                        Login = login,
                        ShortName = shortName,
                        Password = password,
                        Group = groupName
                    };

                    personsToAdd.Add(newPerson);
                    existingLogins.Add(login);

                    string[] validRow = new string[] { lastName, firstName, middleName, login, shortName, password };
                    validRows.Add(validRow);

                    Console.WriteLine($"Строка {rowNumber}: {lastName} {firstName} (логин: {login}, пароль: {password}) добавлен.");
                }

                existingPersons.AddRange(personsToAdd);

                string updatedJson = JsonSerializer.Serialize(existingPersons, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
                File.WriteAllText(filePath, updatedJson, System.Text.Encoding.UTF8);

                Console.WriteLine($"\n✓ Всего добавлено записей: {personsToAdd.Count}");
                Console.WriteLine($"✓ Общее количество записей в файле: {existingPersons.Count}");
                Console.WriteLine($"✓ Файл сохранён: {filePath}\n");

                return (groupName, validRows.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ КРИТИЧЕСКАЯ ОШИБКА: {ex.GetType().Name}");
                Console.WriteLine($"Сообщение: {ex.Message}");
                Console.WriteLine($"Путь к файлу: {filePath}");
                Console.WriteLine($"\nStack Trace:\n{ex.StackTrace}");
                throw;
            }
        }
    }
}