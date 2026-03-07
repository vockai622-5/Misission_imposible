using macros;
using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;
using TFlex.PdmFramework.Resolve;
using System.IO;

namespace ConsoleUsersTest
{
    class Program
    {
        private static readonly Random RandomGenerator = new Random();
        private static string GetValidExcelFilePath()
        {
            while (true)
            {
                Console.WriteLine("Введите путь к Excel файлу (.xlsx или .xls):");
                string input = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(input))
                {
                    Console.WriteLine("Ошибка: путь не может быть пустым.");
                    Console.WriteLine();
                    continue;
                }

                string filePath = input.Trim().Trim('"');

                if (!File.Exists(filePath))
                {
                    Console.WriteLine("Ошибка: файл не найден по указанному пути.");
                    Console.WriteLine();
                    continue;
                }

                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls")
                {
                    Console.WriteLine("Ошибка: файл должен иметь расширение .xlsx или .xls.");
                    Console.WriteLine();
                    continue;
                }

                Console.WriteLine("Файл найден: " + filePath);
                Console.WriteLine();
                return filePath;
            }
        }


            [STAThread]
            static void Main(string[] args)
            {
                var processor = new StudentProcessor();

                string excelFilePath = GetValidExcelFilePath();

                var result = processor.ProcessStudentsFromExcel(excelFilePath);
                string groupName = result.groupName;
                string[][] validData = result.validData;

                AssemblyResolver.Instance.AddDirectory(
                    @"C:\Program Files (x86)\T-FLEX DOCs 17\Program");

                try
                {
                    MainCore(groupName, validData);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Критическая ошибка:");
                    Console.WriteLine(ex);
                }

                Console.WriteLine();
                Console.WriteLine("Нажмите Enter для выхода...");
                Console.ReadLine();
            }

            private static void MainCore(string groupName, string[][] validData)
            {
                using (var connection = ServerConnection.Open())
                {
                    if (connection == null || !connection.IsConnected)
                    {
                        Console.WriteLine("Подключение к серверу отсутствует");
                        return;
                    }

                    Console.WriteLine("Подключение есть");
                    Console.WriteLine("ServerName = " + connection.ServerName);
                    Console.WriteLine("InstanceName = " + connection.InstanceName);
                    Console.WriteLine("CurrentConfiguration = " + connection.CurrentConfiguration);
                    Console.WriteLine("ClientView = " + connection.ClientView);

                    var userReference = new UserReference(connection);

                    Console.WriteLine();
                    Console.WriteLine("Reference = " + userReference.Name);
                    Console.WriteLine("Reference Id = " + userReference.Id);

                    var students = ParseStudents(validData).ToList();

                    if (students.Count == 0)
                    {
                        Console.WriteLine("Не найдено ни одной корректной записи студента.");
                        return;
                    }

                    string finalGroupName = !string.IsNullOrWhiteSpace(groupName)
                        ? groupName
                        : students.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s.Group))?.Group;

                    var testGroup = CreateTopLevelTestGroup(userReference, finalGroupName);
                    if (testGroup == null)
                        return;

                    var createdUsers = new List<User>();

                    foreach (var student in students)
                    {
                        Console.WriteLine(
                            $"Разобрана запись: Фамилия='{student.LastName}', Имя='{student.FirstName}', Отчество='{student.MiddleName}', Логин='{student.Login}', Группа='{student.Group}'");

                        if (string.IsNullOrWhiteSpace(student.LastName) &&
                            string.IsNullOrWhiteSpace(student.FirstName) &&
                            string.IsNullOrWhiteSpace(student.MiddleName))
                        {
                            Console.WriteLine("Пропуск строки: пустое ФИО.");
                            continue;
                        }

                        var user = CreateTestEmployee(userReference, testGroup, student);

                        if (user != null)
                            createdUsers.Add(user);
                    }

                    if (createdUsers.Count > 0)
                    {
                        string workspaceFolderName = !string.IsNullOrWhiteSpace(finalGroupName)
                            ? finalGroupName
                            : "Test_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                        CreateStudentWorkspacesInFiles(connection, workspaceFolderName, createdUsers);
                    }
                    else
                    {
                        Console.WriteLine("Не удалось создать ни одного пользователя.");
                    }
                }
            }

            private static List<StudentImportRow> ParseStudents(string[][] validData)
            {
                var result = new List<StudentImportRow>();

                if (validData == null || validData.Length == 0)
                    return result;

                for (int i = 0; i < validData.Length; i++)
                {
                    var row = validData[i];
                    if (row == null || row.Length == 0)
                    {
                        Console.WriteLine($"Строка {i}: пустая.");
                        continue;
                    }

                    var student = ExtractStudentRow(row);

                    if (student == null)
                    {
                        Console.WriteLine($"Строка {i}: не удалось разобрать.");
                        continue;
                    }

                    result.Add(student);
                }

                return result;
            }

            private static StudentImportRow ExtractStudentRow(string[] row)
            {
                if (row == null)
                    return null;

                return new StudentImportRow
                {
                    LastName = GetValue(row, 0),
                    FirstName = GetValue(row, 1),
                    MiddleName = GetValue(row, 2),
                    Login = GetValue(row, 3),
                    ShortName = GetValue(row, 4),
                    Password = GetValue(row, 5),
                    Group = GetValue(row, 6)
                };
            }

            private static string GetValue(string[] row, int index)
            {
                if (row == null)
                    return string.Empty;

                if (index < 0 || index >= row.Length)
                    return string.Empty;

                return row[index]?.Trim() ?? string.Empty;
            }

            private static UserReferenceObject CreateTopLevelTestGroup(UserReference userReference, string groupName)
            {
                Console.WriteLine();
                Console.WriteLine("=== СОЗДАНИЕ ВЕРХНЕУРОВНЕВОЙ ГРУППЫ ===");

                var groupType = userReference.Classes.GroupBaseType;
                Console.WriteLine("Group type = " + groupType?.Name);

                string testName = string.IsNullOrWhiteSpace(groupName)
                    ? "_API_WRITE_TEST_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    : groupName;

                try
                {
                    var newGroup = userReference.CreateReferenceObject(groupType);

                    Console.WriteLine("Объект группы создан в памяти:");
                    PrintUserReferenceObjectShort(newGroup);

                    newGroup.FullName.Value = testName;

                    if (newGroup.Description != null)
                        newGroup.Description.Value = "Группа, созданная из Excel через внешнее приложение";

                    CommitObject(newGroup);

                    Console.WriteLine();
                    Console.WriteLine("После коммита группы:");
                    PrintUserReferenceObjectShort(newGroup);

                    Console.WriteLine("УСПЕХ: группа создана");
                    return newGroup;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при создании группы:");
                    Console.WriteLine(ex);
                    return null;
                }
            }

            private static User CreateTestEmployee(
                UserReference userReference,
                ReferenceObject parentGroup,
                StudentImportRow student)
            {
                Console.WriteLine();
                Console.WriteLine("=== СОЗДАНИЕ СОТРУДНИКА ===");
                Console.WriteLine("ParentGroupId = " + parentGroup.Id);

                string fullName = BuildFullName(student.LastName, student.FirstName, student.MiddleName);
                string shortName = !string.IsNullOrWhiteSpace(student.ShortName)
                    ? student.ShortName
                    : BuildShortName(student.LastName, student.FirstName, student.MiddleName);

                string login = !string.IsNullOrWhiteSpace(student.Login)
                    ? student.Login
                    : BuildLogin(student.LastName, student.FirstName, student.MiddleName);

                Console.WriteLine("ФИО = " + fullName);
                Console.WriteLine("ShortName = " + shortName);
                Console.WriteLine("Login = " + login);

                try
                {
                    var userType = userReference.Classes.EmployerType;
                    Console.WriteLine("User type = " + userType?.Name);

                    var newUser = (User)userReference.CreateReferenceObject(parentGroup, userType);

                    Console.WriteLine("Пользователь создан в памяти:");
                    PrintUserReferenceObjectShort(newUser);

                    newUser.FullName.Value = fullName;
                    newUser.LastName.Value = student.LastName ?? string.Empty;
                    newUser.FirstName.Value = student.FirstName ?? string.Empty;
                    newUser.Patronymic.Value = student.MiddleName ?? string.Empty;
                    newUser.ShortName.Value = shortName;
                    newUser.Login.Value = login;

                    if (newUser.Description != null)
                        newUser.Description.Value = "Пользователь, созданный из Excel через внешнее приложение";

                    CommitObject(newUser);

                    Console.WriteLine();
                    Console.WriteLine("После коммита сотрудника:");
                    PrintUserReferenceObjectShort(newUser);

                    Console.WriteLine("УСПЕХ: сотрудник создан");
                    return newUser;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при создании сотрудника:");
                    Console.WriteLine(ex);
                    return null;
                }
            }

            private static void CreateStudentWorkspacesInFiles(
                ServerConnection connection,
                string groupFolderName,
                IEnumerable<User> users)
            {
                Console.WriteLine();
                Console.WriteLine("=== СОЗДАНИЕ РАБОЧИХ ПАПОК В СПРАВОЧНИКЕ 'ФАЙЛЫ' ===");
                Console.WriteLine("Папка группы = " + groupFolderName);

                try
                {
                    var fileReference = new FileReference(connection);

                    var studentsRoot = FindStudentsFolder(fileReference);
                    if (studentsRoot == null)
                    {
                        Console.WriteLine("Не найдена папка 'Студенты' в справочнике 'Файлы'.");
                        return;
                    }

                    Console.WriteLine("Папка 'Студенты' найдена.");

                    string safeGroupFolderName = MakeSafeFolderName(
                        groupFolderName,
                        "Group_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));

                    FolderObject groupFolder = studentsRoot.CreateFolder(
                        "Тестовая папка группы " + safeGroupFolderName,
                        safeGroupFolderName);

                    var objectsToCheckIn = new List<DesktopObject> { groupFolder };
                    var usedFolderNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var user in users)
                    {
                        string sourceShortName = user.ShortName?.Value;
                        string safeStudentFolderName = MakeSafeFolderName(
                            sourceShortName,
                            "Student_" + user.Id);

                        safeStudentFolderName = MakeUniqueFolderName(safeStudentFolderName, usedFolderNames);

                        FolderObject studentFolder = groupFolder.CreateFolder(
                            "Рабочая папка пользователя " + (sourceShortName ?? safeStudentFolderName),
                            safeStudentFolderName);

                        objectsToCheckIn.Add(studentFolder);
                        Console.WriteLine("Создана папка пользователя: " + safeStudentFolderName);
                    }

                    Desktop.CheckIn(objectsToCheckIn, "Созданы рабочие папки пользователей", false);

                    Console.WriteLine("УСПЕХ: рабочие папки в справочнике 'Файлы' созданы");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при создании рабочих папок в справочнике 'Файлы':");
                    Console.WriteLine(ex);
                }
            }

            private static string MakeSafeFolderName(string value, string fallback)
            {
                string result = value ?? string.Empty;

                result = result.Trim();

                if (string.IsNullOrWhiteSpace(result))
                    result = fallback ?? "Folder";

                var invalidChars = Path.GetInvalidFileNameChars();
                var sanitizedChars = new List<char>(result.Length);

                foreach (char ch in result)
                {
                    if (Array.IndexOf(invalidChars, ch) >= 0 || ch == '/' || ch == '\\' || ch == ':' || ch == '*' || ch == '?' || ch == '"' || ch == '<' || ch == '>' || ch == '|')
                        sanitizedChars.Add('_');
                    else if (char.IsControl(ch))
                        sanitizedChars.Add('_');
                    else
                        sanitizedChars.Add(ch);
                }

                result = new string(sanitizedChars.ToArray());

                while (result.Contains("  "))
                    result = result.Replace("  ", " ");

                result = result.Trim();
                result = result.TrimEnd('.', ' ');

                if (string.IsNullOrWhiteSpace(result))
                    result = fallback ?? "Folder";

                result = result.TrimEnd('.', ' ');

                if (string.IsNullOrWhiteSpace(result))
                    result = "Folder";

                const int maxLength = 120;
                if (result.Length > maxLength)
                    result = result.Substring(0, maxLength).TrimEnd('.', ' ');

                if (string.IsNullOrWhiteSpace(result))
                    result = "Folder";

                return result;
            }

            private static string MakeUniqueFolderName(string baseName, HashSet<string> usedNames)
            {
                if (usedNames == null)
                    return baseName;

                string candidate = baseName;
                int index = 2;

                while (usedNames.Contains(candidate))
                {
                    string suffix = " (" + index + ")";
                    int maxBaseLength = Math.Max(1, 120 - suffix.Length);
                    string truncatedBase = baseName.Length > maxBaseLength
                        ? baseName.Substring(0, maxBaseLength).TrimEnd('.', ' ')
                        : baseName;

                    if (string.IsNullOrWhiteSpace(truncatedBase))
                        truncatedBase = "Folder";

                    candidate = truncatedBase + suffix;
                    candidate = candidate.TrimEnd('.', ' ');
                    index++;
                }

                usedNames.Add(candidate);
                return candidate;
            }

            private static FolderObject FindStudentsFolder(FileReference fileReference)
            {
                var folder =
                    TryFindFolderByRelativePath(fileReference, @"Файлы\Студенты") ??
                    TryFindFolderByRelativePath(fileReference, @"Файлы\Файлы\Студенты") ??
                    TryFindFolderByRelativePath(fileReference, "Студенты") ??
                    TryFindFolderByRelativePath(fileReference, @"\Файлы\Студенты") ??
                    TryFindFolderByRelativePath(fileReference, @"\Файлы\Файлы\Студенты") ??
                    TryFindFolderByRelativePath(fileReference, @"\Студенты");

                if (folder != null)
                    return folder;

                var filesRoot =
                    TryFindFolderByRelativePath(fileReference, "Файлы") ??
                    TryFindFolderByRelativePath(fileReference, @"\Файлы") ??
                    TryFindFolderByRelativePath(fileReference, @"Файлы\Файлы") ??
                    TryFindFolderByRelativePath(fileReference, @"\Файлы\Файлы");

                if (filesRoot == null)
                    return null;

                return FindChildFolderByName(filesRoot, "Студенты");
            }

            private static FolderObject TryFindFolderByRelativePath(FileReference fileReference, string path)
            {
                try
                {
                    return fileReference.FindByRelativePath(path) as FolderObject;
                }
                catch
                {
                    return null;
                }
            }

            private static FolderObject FindChildFolderByName(FolderObject parentFolder, string folderName)
            {
                if (parentFolder == null || string.IsNullOrWhiteSpace(folderName))
                    return null;

                try
                {
                    parentFolder.Children.Load();

                    foreach (ReferenceObject child in parentFolder.Children)
                    {
                        var childFolder = child as FolderObject;
                        if (childFolder == null)
                            continue;

                        string childName = GetFolderName(childFolder);

                        if (string.Equals(childName, folderName, StringComparison.OrdinalIgnoreCase))
                            return childFolder;

                        var nested = FindChildFolderByName(childFolder, folderName);
                        if (nested != null)
                            return nested;
                    }
                }
                catch
                {
                }

                return null;
            }

            private static string GetFolderName(FolderObject folder)
            {
                if (folder == null)
                    return string.Empty;

                try
                {
                    var nameProp = folder.GetType().GetProperty("Name");
                    if (nameProp == null)
                        return string.Empty;

                    object rawValue = nameProp.GetValue(folder);
                    if (rawValue == null)
                        return string.Empty;

                    var valueProp = rawValue.GetType().GetProperty("Value");
                    if (valueProp != null)
                    {
                        object value = valueProp.GetValue(rawValue);
                        return value?.ToString() ?? string.Empty;
                    }

                    return rawValue.ToString();
                }
                catch
                {
                    return string.Empty;
                }
            }

            private static void CommitObject(UserReferenceObject obj)
            {
                bool ok = obj.ApplyChanges();
                if (!ok)
                    throw new InvalidOperationException("Не удалось применить изменения объекта.");
            }

            private static void PrintUserReferenceObjectShort(UserReferenceObject obj)
            {
                if (obj == null)
                {
                    Console.WriteLine("obj = null");
                    return;
                }

                Console.WriteLine(
                    "Type = " + obj.GetType().Name +
                    " | Id = " + obj.Id +
                    " | Guid = " + obj.Guid +
                    " | Class = " + obj.Class?.Name +
                    " | IsGroup = " + obj.IsGroup +
                    " | IsUser = " + obj.IsUser +
                    " | IsNew = " + obj.IsNew +
                    " | IsModified = " + obj.IsModified +
                    " | IsChanged = " + obj.IsChanged);
            }

            private static string BuildFullName(string lastName, string firstName, string middleName)
            {
                var parts = new List<string>();

                if (!string.IsNullOrWhiteSpace(lastName))
                    parts.Add(lastName.Trim());

                if (!string.IsNullOrWhiteSpace(firstName))
                    parts.Add(firstName.Trim());

                if (!string.IsNullOrWhiteSpace(middleName))
                    parts.Add(middleName.Trim());

                return string.Join(" ", parts);
            }

            private static string BuildShortName(string lastName, string firstName, string middleName)
            {
                string firstInitial = GetInitial(firstName);
                string middleInitial = GetInitial(middleName);

                string result = (lastName ?? string.Empty).Trim();

                if (!string.IsNullOrWhiteSpace(firstInitial))
                    result += " " + firstInitial + ".";

                if (!string.IsNullOrWhiteSpace(middleInitial))
                    result += " " + middleInitial + ".";

                return result.Trim();
            }

            private static string BuildLogin(string lastName, string firstName, string middleName)
            {
                string surnamePart = NormalizeLoginPart(Transliterate(lastName));
                string firstInitial = NormalizeLoginPart(Transliterate(GetInitial(firstName)));
                string middleInitial = NormalizeLoginPart(Transliterate(GetInitial(middleName)));
                string randomSuffix = GenerateRandomSuffix(4);

                if (string.IsNullOrWhiteSpace(surnamePart))
                    surnamePart = "user";

                if (!string.IsNullOrWhiteSpace(firstInitial) && !string.IsNullOrWhiteSpace(middleInitial))
                    return $"{surnamePart}_{firstInitial}_{middleInitial}_{randomSuffix}";

                if (!string.IsNullOrWhiteSpace(firstInitial))
                    return $"{surnamePart}_{firstInitial}_{randomSuffix}";

                return $"{surnamePart}_{randomSuffix}";
            }

            private static string GenerateRandomSuffix(int digits)
            {
                if (digits <= 0)
                    digits = 4;

                int min = (int)Math.Pow(10, digits - 1);
                int max = (int)Math.Pow(10, digits) - 1;

                lock (RandomGenerator)
                {
                    return RandomGenerator.Next(min, max + 1).ToString();
                }
            }

            private static string GetInitial(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return string.Empty;

                value = value.Trim();
                return value.Length > 0 ? value.Substring(0, 1) : string.Empty;
            }

            private static string NormalizeLoginPart(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return string.Empty;

                value = value.ToLowerInvariant();
                var chars = new List<char>();

                foreach (char ch in value)
                {
                    if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9'))
                        chars.Add(ch);
                }

                return new string(chars.ToArray());
            }

            private static string Transliterate(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    return string.Empty;

                var map = new Dictionary<char, string>
                {
                    ['а'] = "a",
                    ['б'] = "b",
                    ['в'] = "v",
                    ['г'] = "g",
                    ['д'] = "d",
                    ['е'] = "e",
                    ['ё'] = "e",
                    ['ж'] = "zh",
                    ['з'] = "z",
                    ['и'] = "i",
                    ['й'] = "y",
                    ['к'] = "k",
                    ['л'] = "l",
                    ['м'] = "m",
                    ['н'] = "n",
                    ['о'] = "o",
                    ['п'] = "p",
                    ['р'] = "r",
                    ['с'] = "s",
                    ['т'] = "t",
                    ['у'] = "u",
                    ['ф'] = "f",
                    ['х'] = "kh",
                    ['ц'] = "ts",
                    ['ч'] = "ch",
                    ['ш'] = "sh",
                    ['щ'] = "sch",
                    ['ъ'] = "",
                    ['ы'] = "y",
                    ['ь'] = "",
                    ['э'] = "e",
                    ['ю'] = "yu",
                    ['я'] = "ya"
                };

                value = value.Trim().ToLowerInvariant();
                var result = new List<string>();

                foreach (char ch in value)
                {
                    if (map.TryGetValue(ch, out string mapped))
                        result.Add(mapped);
                    else
                        result.Add(ch.ToString());
                }

                return string.Join("", result);
            }

            private class StudentImportRow
            {
                public string LastName { get; set; }
                public string FirstName { get; set; }
                public string MiddleName { get; set; }
                public string Login { get; set; }
                public string ShortName { get; set; }
                public string Password { get; set; }
                public string Group { get; set; }
            }
        }
    }