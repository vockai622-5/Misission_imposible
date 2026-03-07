using System;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.References.Users;
using TFlex.PdmFramework.Resolve;

namespace TFlexDocsUsers
{
    /// <summary>
    /// Пример получения списка пользователей из T-FLEX DOCs 17.5.4.0
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // Настройки подключения — измените значения под ваш сервер
            string serverAddress = "localhost";
            string login = "admin";
            string password = "admin";

            Console.WriteLine("=== T-FLEX DOCs — Получение списка пользователей ===");
            Console.WriteLine($"Подключение к серверу: {serverAddress}");

            ServerConnection connection = null;

            try
            {
                // Открытие соединения с сервером T-FLEX DOCs
                connection = ServerConnection.Open(serverAddress, login, password);
                Console.WriteLine("Подключение установлено успешно.");

                // Получение справочника пользователей
                var userReference = new UserReference(connection);

                // --- Получение всех пользователей ---
                Console.WriteLine("\n--- Все пользователи ---");
                var users = userReference.GetAllUsers();
                if (users == null || users.Count == 0)
                {
                    Console.WriteLine("Пользователи не найдены.");
                }
                else
                {
                    foreach (var user in users)
                    {
                        Console.WriteLine($"  Логин: {user.Login,-20}  Имя: {user.Name}");
                    }
                    Console.WriteLine($"Итого пользователей: {users.Count}");
                }

                // --- Получение всех групп пользователей ---
                Console.WriteLine("\n--- Группы пользователей ---");
                var groups = userReference.GetAllUsersGroup();
                if (groups == null || groups.Count == 0)
                {
                    Console.WriteLine("Группы не найдены.");
                }
                else
                {
                    foreach (var group in groups)
                    {
                        Console.WriteLine($"  Группа: {group.Name}");
                    }
                    Console.WriteLine($"Итого групп: {groups.Count}");
                }
            }
            catch (ResolveException ex)
            {
                // Ошибка при разрешении адреса сервера
                Console.Error.WriteLine($"[Ошибка подключения] Не удалось найти сервер '{serverAddress}': {ex.Message}");
                Environment.Exit(1);
            }
            catch (AuthenticationException ex)
            {
                // Неверные учётные данные
                Console.Error.WriteLine($"[Ошибка аутентификации] Проверьте логин и пароль: {ex.Message}");
                Environment.Exit(1);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Неожиданная ошибка] {ex.GetType().Name}: {ex.Message}");
                Environment.Exit(1);
            }
            finally
            {
                // Всегда закрываем соединение
                connection?.Close();
                Console.WriteLine("\nСоединение закрыто.");
            }
        }
    }
}
