using System;
using System.Linq;
using System.Reflection;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.PdmFramework.Resolve;

namespace ConsoleUsersTest
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
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

                CreateTopLevelTestGroup(userReference);
            }
        }

        private static void CreateTopLevelTestGroup(UserReference userReference)
        {
            Console.WriteLine();
            Console.WriteLine("=== ТЕСТ ЗАПИСИ: создание верхнеуровневой группы ===");

            var groupType = userReference.Classes.GroupBaseType;
            Console.WriteLine("Group type = " + groupType?.Name);

            string testName = "_API_WRITE_TEST_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

            try
            {
                var newGroup = userReference.CreateReferenceObject(groupType);

                Console.WriteLine("Объект создан в памяти:");
                PrintObjectShort(newGroup);

                // Для нового объекта BeginChanges НЕ нужен
                newGroup.FullName.Value = testName;

                if (newGroup.Description != null)
                    newGroup.Description.Value = "Тестовая группа, создана из внешнего приложения";

                Console.WriteLine("Поля заполнены:");
                Console.WriteLine("FullName = " + newGroup.FullName.Value);
                Console.WriteLine("Description = " + (newGroup.Description == null ? "<null>" : newGroup.Description.Value));

                Console.WriteLine();
                Console.WriteLine("=== Диагностика коммита ===");
                Console.WriteLine("SaveSet = " + (newGroup.SaveSet == null ? "null" : newGroup.SaveSet.GetType().FullName));

                DumpCommitLikeMethods(newGroup.GetType(), "Runtime object");
                DumpCommitLikeMethods(typeof(ReferenceObject), "ReferenceObject");
                DumpCommitLikeMethods(typeof(ReferenceObjectSaveSet), "ReferenceObjectSaveSet");
                DumpCommitLikeMethods(typeof(SaveSetEditSession), "SaveSetEditSession");
                DumpCommitLikeMethods(typeof(EditSession), "EditSession");

                CommitNewObject(newGroup);

                Console.WriteLine();
                Console.WriteLine("После коммита:");
                PrintObjectShort(newGroup);

                Console.WriteLine("УСПЕХ: тестовая группа создана");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при создании группы:");
                Console.WriteLine(ex);
            }
        }

        private static void CommitNewObject(UserReferenceObject obj)
        {
            // 1. Сначала пробуем через SaveSet — это теперь основной кандидат
            if (obj.SaveSet != null)
            {
                Console.WriteLine();
                Console.WriteLine("Пробую коммит через SaveSet...");

                if (TryKnownCommitMethods(obj.SaveSet))
                    return;

                if (TryCommitViaSessionFactory(obj.SaveSet))
                    return;
            }

            // 2. Потом уже пробуем на самом объекте
            Console.WriteLine();
            Console.WriteLine("Пробую коммит на самом объекте...");

            if (TryKnownCommitMethods(obj))
                return;

            if (TryCommitViaSessionFactory(obj))
                return;

            throw new NotSupportedException(
                "Не найден рабочий способ коммита. " +
                "Нужен вывод блока 'Диагностика коммита'.");
        }

        private static bool TryKnownCommitMethods(object target)
        {
            string[] methodNames =
            {
                "Save",
                "Apply",
                "ApplyChanges",
                "Commit",
                "CheckIn",
                "EndEdit",
                "EndChanges"
            };

            var type = target.GetType();

            foreach (var name in methodNames)
            {
                var method = type.GetMethod(name, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (method == null)
                    continue;

                Console.WriteLine($"Использую {type.Name}.{name}()");
                method.Invoke(target, null);
                return true;
            }

            return false;
        }

        private static bool TryCommitViaSessionFactory(object target)
        {
            string[] factoryNames =
            {
                "BeginEdit",
                "CreateEditSession",
                "BeginSession",
                "Edit"
            };

            var type = target.GetType();

            foreach (var factoryName in factoryNames)
            {
                var factory = type.GetMethod(factoryName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (factory == null)
                    continue;

                Console.WriteLine($"Пробую через {type.Name}.{factoryName}()");

                object session = null;
                try
                {
                    session = factory.Invoke(target, null);
                    if (session == null)
                        continue;

                    Console.WriteLine("Session type = " + session.GetType().FullName);

                    if (TryKnownCommitMethods(session))
                    {
                        TryDispose(session);
                        return true;
                    }

                    TryDispose(session);
                }
                catch
                {
                    TryDispose(session);
                    throw;
                }
            }

            return false;
        }

        private static void TryDispose(object obj)
        {
            if (obj is IDisposable d)
                d.Dispose();
        }

        private static void DumpCommitLikeMethods(Type type, string title)
        {
            Console.WriteLine();
            Console.WriteLine("=== " + title + " ===");
            Console.WriteLine(type.FullName);

            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m =>
                    m.Name.IndexOf("save", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("apply", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("commit", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("check", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("edit", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("change", StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(m => m.Name);

            foreach (var m in methods)
            {
                var ps = string.Join(", ", m.GetParameters()
                    .Select(p => p.ParameterType.Name + " " + p.Name));

                Console.WriteLine($"{m.ReturnType.Name} {m.Name}({ps})");
            }
        }

        private static void PrintObjectShort(UserReferenceObject obj)
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
    }
}