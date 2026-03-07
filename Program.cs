using macros;
using System;
using System.Collections.Generic;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;
using TFlex.PdmFramework.Resolve;

namespace ConsoleUsersTest
{
    class Program
    {
        private static readonly Random RandomGenerator = new Random();

        [STAThread]
        static void Main(string[] args)
        {
            var processor = new StudentProcessor();

            string excelFilePath = @"C:\Users\astaf\Desktop\data.xlsx";

            // Вызываем одну функцию - она делает всё
            var (groupName, validData) = processor.ProcessStudentsFromExcel(excelFilePath);


            AssemblyResolver.Instance.AddDirectory(
                @"C:\Program Files (x86)\T-FLEX DOCs 17\Program");

            try
            {
                MainCore();
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

        private static void MainCore()
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

                var testGroup = CreateTopLevelTestGroup(userReference);
                if (testGroup == null)
                    return;

                var createdUsers = new List<User>();

                var user1 = CreateTestEmployee(userReference, testGroup, "Шестаков", "Михаил", "Викторович");
                if (user1 != null)
                    createdUsers.Add(user1);

                var user2 = CreateTestEmployee(userReference, testGroup, "Корнаухов", "Иван", "Сергеевич");
                if (user2 != null)
                    createdUsers.Add(user2);

                if (createdUsers.Count > 0)
                {
                    string workspaceFolderName = "Test_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    CreateStudentWorkspacesInFiles(connection, workspaceFolderName, createdUsers);
                }
            }
        }

        private static UserReferenceObject CreateTopLevelTestGroup(UserReference userReference)
        {
            Console.WriteLine();
            Console.WriteLine("=== СОЗДАНИЕ ВЕРХНЕУРОВНЕВОЙ ГРУППЫ ===");

            var groupType = userReference.Classes.GroupBaseType;
            Console.WriteLine("Group type = " + groupType?.Name);

            string testName = "_API_WRITE_TEST_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

            try
            {
                var newGroup = userReference.CreateReferenceObject(groupType);

                Console.WriteLine("Объект группы создан в памяти:");
                PrintUserReferenceObjectShort(newGroup);

                newGroup.FullName.Value = testName;

                if (newGroup.Description != null)
                    newGroup.Description.Value = "Тестовая группа, создана из внешнего приложения";

                CommitObject(newGroup);

                Console.WriteLine();
                Console.WriteLine("После коммита группы:");
                PrintUserReferenceObjectShort(newGroup);

                Console.WriteLine("УСПЕХ: тестовая группа создана");
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
            string lastName,
            string firstName,
            string middleName)
        {
            Console.WriteLine();
            Console.WriteLine("=== СОЗДАНИЕ СОТРУДНИКА ===");
            Console.WriteLine("ParentGroupId = " + parentGroup.Id);

            string fullName = BuildFullName(lastName, firstName, middleName);
            string shortName = BuildShortName(lastName, firstName, middleName);
            string login = BuildLogin(lastName, firstName, middleName);

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
                newUser.LastName.Value = lastName ?? string.Empty;
                newUser.FirstName.Value = firstName ?? string.Empty;
                newUser.Patronymic.Value = middleName ?? string.Empty;
                newUser.ShortName.Value = shortName;
                newUser.Login.Value = login;

                if (newUser.Description != null)
                    newUser.Description.Value = "Тестовый сотрудник из внешнего приложения";

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

                // Порядок аргументов: description, name
                FolderObject groupFolder = studentsRoot.CreateFolder(
                    "Тестовая папка группы " + groupFolderName,
                    groupFolderName);

                var objectsToCheckIn = new List<DesktopObject> { groupFolder };

                foreach (var user in users)
                {
                    string shortName = user.ShortName?.Value;

                    if (string.IsNullOrWhiteSpace(shortName))
                    {
                        Console.WriteLine("Пропуск пользователя: пустое короткое имя.");
                        continue;
                    }

                    FolderObject studentFolder = groupFolder.CreateFolder(
                        "Рабочая папка пользователя " + shortName,
                        shortName);

                    objectsToCheckIn.Add(studentFolder);
                    Console.WriteLine("Создана папка пользователя: " + shortName);
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

        private static FolderObject FindStudentsFolder(FileReference fileReference)
        {
            // Сначала пробуем самые вероятные относительные пути
            var folder =
                TryFindFolderByRelativePath(fileReference, @"Файлы\Студенты") ??
                TryFindFolderByRelativePath(fileReference, @"Файлы\Файлы\Студенты") ??
                TryFindFolderByRelativePath(fileReference, "Студенты") ??
                TryFindFolderByRelativePath(fileReference, @"\Файлы\Студенты") ??
                TryFindFolderByRelativePath(fileReference, @"\Файлы\Файлы\Студенты") ??
                TryFindFolderByRelativePath(fileReference, @"\Студенты");

            if (folder != null)
                return folder;

            // Если прямой путь не сработал, пытаемся найти корневую папку "Файлы"
            var filesRoot =
                TryFindFolderByRelativePath(fileReference, "Файлы") ??
                TryFindFolderByRelativePath(fileReference, @"\Файлы") ??
                TryFindFolderByRelativePath(fileReference, @"Файлы\Файлы") ??
                TryFindFolderByRelativePath(fileReference, @"\Файлы\Файлы");

            if (filesRoot == null)
                return null;

            // Ищем "Студенты" рекурсивно внутри найденного корня
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
    }
}