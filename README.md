# T-FLEX DOCs — Получение списка пользователей (C#)

Рабочий пример кода на **C#** для получения списка всех пользователей и групп из **T-FLEX DOCs 17.5.4.0** с помощью официального SDK.

---

## Содержание

- [Требования](#требования)
- [Установка зависимостей](#установка-зависимостей)
- [Настройка подключения](#настройка-подключения)
- [Сборка и запуск](#сборка-и-запуск)
- [Примеры использования API](#примеры-использования-api)
- [Обработка ошибок](#обработка-ошибок)
- [Ссылки на документацию](#ссылки-на-документацию)

---

## Требования

| Компонент | Версия |
|-----------|--------|
| T-FLEX DOCs | 17.5.4.0 |
| .NET Framework | 4.8 |
| Visual Studio | 2019 / 2022 |
| ОС | Windows 10/11 или Windows Server 2016+ |

---

## Установка зависимостей

T-FLEX DOCs SDK **не распространяется через NuGet**. Необходимые сборки (DLL) входят в состав установки T-FLEX DOCs.

### Шаги

1. Установите **T-FLEX DOCs 17.5.4.0** (сервер или клиент) на рабочей машине разработчика.  
   Путь по умолчанию: `C:\Program Files\T-FLEX DOCs 17\`

2. Скопируйте следующие файлы в папку `TFlexDocsUsers\libs\` этого проекта:

   | Файл DLL | Пространство имён |
   |---|---|
   | `TFlex.DOCs.Common.dll` | `TFlex.DOCs.Common` |
   | `TFlex.DOCs.Model.dll` | `TFlex.DOCs.Model.References.Users` |
   | `TFlex.PdmFramework.dll` | `TFlex.PdmFramework.Resolve` |

   ```powershell
   # Пример команды PowerShell (путь установки по умолчанию)
   $src = "C:\Program Files\T-FLEX DOCs 17"
   $dst = ".\TFlexDocsUsers\libs"
   New-Item -ItemType Directory -Force -Path $dst
   Copy-Item "$src\TFlex.DOCs.Common.dll"   $dst
   Copy-Item "$src\TFlex.DOCs.Model.dll"    $dst
   Copy-Item "$src\TFlex.PdmFramework.dll"  $dst
   ```

3. Откройте решение в Visual Studio или соберите через .NET CLI:

   ```bash
   cd TFlexDocsUsers
   dotnet build
   ```

---

## Настройка подключения

Откройте `TFlexDocsUsers/Program.cs` и измените три строки в начале метода `Main`:

```csharp
string serverAddress = "localhost";   // Имя или IP-адрес сервера T-FLEX DOCs
string login         = "admin";       // Логин пользователя
string password      = "admin";       // Пароль пользователя
```

> **Совет.** Для продуктивных сред рекомендуется читать учётные данные из переменных окружения или конфигурационного файла, а не хранить их в исходном коде.

---

## Сборка и запуск

```bash
cd TFlexDocsUsers
dotnet build
dotnet run
```

Пример вывода:

```
=== T-FLEX DOCs — Получение списка пользователей ===
Подключение к серверу: localhost
Подключение установлено успешно.

--- Все пользователи ---
  Логин: admin                Имя: Администратор
  Логин: ivanov               Имя: Иванов Иван Иванович
  Логин: petrov               Имя: Петров Пётр Петрович
Итого пользователей: 3

--- Группы пользователей ---
  Группа: Администраторы
  Группа: Конструкторы
Итого групп: 2

Соединение закрыто.
```

---

## Примеры использования API

### Подключение к серверу

```csharp
using TFlex.DOCs.Common;

// Открытие соединения
ServerConnection connection = ServerConnection.Open(serverAddress, login, password);

// ... работа с API ...

// Закрытие соединения (всегда в блоке finally)
connection.Close();
```

### Получение всех пользователей

```csharp
using TFlex.DOCs.Model.References.Users;

var userReference = new UserReference(connection);

// Все пользователи системы
var users = userReference.GetAllUsers();
foreach (var user in users)
{
    Console.WriteLine($"Логин: {user.Login}, Имя: {user.Name}");
}
```

### Получение групп пользователей

```csharp
// Все группы пользователей
var groups = userReference.GetAllUsersGroup();
foreach (var group in groups)
{
    Console.WriteLine($"Группа: {group.Name}");
}
```

### Поиск пользователя по логину

```csharp
// Найти конкретного пользователя по логину
var user = userReference.Find(u => u.Login == "ivanov");
if (user != null)
{
    Console.WriteLine($"Найден: {user.Name}, Email: {user.Email}");
}
```

### Полный шаблон с обработкой ошибок

```csharp
using System;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model.References.Users;
using TFlex.PdmFramework.Resolve;

ServerConnection connection = null;
try
{
    connection = ServerConnection.Open("localhost", "admin", "admin");
    var userRef = new UserReference(connection);

    var users = userRef.GetAllUsers();
    foreach (var user in users)
        Console.WriteLine($"{user.Login} — {user.Name}");
}
catch (ResolveException ex)
{
    Console.Error.WriteLine($"Сервер не найден: {ex.Message}");
}
catch (AuthenticationException ex)
{
    Console.Error.WriteLine($"Ошибка аутентификации: {ex.Message}");
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Ошибка: {ex.Message}");
}
finally
{
    connection?.Close();
}
```

---

## Обработка ошибок

| Исключение | Причина | Решение |
|---|---|---|
| `ResolveException` | Сервер недоступен или не найден по адресу | Проверьте `serverAddress` и доступность сети |
| `AuthenticationException` | Неверный логин или пароль | Проверьте учётные данные пользователя |
| `Exception` | Любая другая ошибка SDK | Смотрите `ex.Message` для подробностей |

---

## Ссылки на документацию

- [Официальный сайт T-FLEX DOCs](https://www.tflex.ru/products/docs/)
- [Портал документации Top Systems](https://kb.tflex.ru/)
- [Справка по SDK T-FLEX DOCs (раздел «Разработка»)](https://kb.tflex.ru/docs/api)
- [Форум разработчиков T-FLEX](https://forum.tflex.ru/viewforum.php?f=10)

> **Примечание.** Подробная документация по API (описание классов `ServerConnection`, `UserReference` и всех методов) включена в поставку T-FLEX DOCs SDK и доступна через **Object Browser** в Visual Studio после подключения сборок.
