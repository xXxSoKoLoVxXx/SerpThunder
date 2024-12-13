using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using System.Security.Cryptography;
using SerpThunder;

class Program
{
    private static readonly string BotToken = "7672299608:AAETZ0HDDPKPMPH16TqnCS96cuLJmBZ4wjc";
    private static TelegramBotClient botClient;
    private static string currentFilePath;
    private static readonly string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\Schedule\");
    static System.Threading.Timer updateTimer;
    private static Sub subscriptionManager;
    private static Format formatManager;

    // Хэш предыдущего файла
    private static string previousFileHash = null;

    static async Task Main(string[] args)
    {
        string formatFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\formats.txt");
        formatManager = new Format(formatFilePath);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string subscriptionFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\subscriptions.txt");
        subscriptionManager = new Sub(subscriptionFilePath);
        subscriptionManager.LoadSubscriptions();

        currentFilePath = await DownloadAndSaveFileAsync("https://serp-koll.ru/images/ep/k1/rasp1.xlsx", directoryPath);
        previousFileHash = ComputeFileHash(currentFilePath);

        botClient = new TelegramBotClient(BotToken);
        Console.WriteLine("Бот запущен...");

        using var cts = new CancellationTokenSource();
        var receiverOptions = new ReceiverOptions
        {
            AllowedUpdates = Array.Empty<UpdateType>()
        };

        botClient.StartReceiving(
            HandleUpdateAsync,
            HandleErrorAsync,
            receiverOptions,
            cancellationToken: cts.Token
        );

        // Таймер для проверки обновлений
        updateTimer = new System.Threading.Timer(async _ =>
        {
            await UpdateScheduleAndNotify();
        }, null, TimeSpan.FromMinutes(10), TimeSpan.FromMinutes(10));

        // Ждём завершения через сигнал отмены
        AppDomain.CurrentDomain.ProcessExit += (s, e) =>
        {
            Console.WriteLine("Бот завершает работу...");
            cts.Cancel(); // Отмена всех задач
        };

        try
        {
            await Task.Delay(Timeout.Infinite, cts.Token); // Ожидание завершения
        }
        catch (TaskCanceledException)
        {
            // Нормальное завершение
        }
        finally
        {
            Console.WriteLine("Бот остановлен.");
        }
    }

    private static async Task UpdateScheduleAndNotify()
    {
        Console.WriteLine("Проверка обновления расписания...");

        string tempPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\temp_schedule");        

        await DownloadAndSaveFileAsync("https://serp-koll.ru/images/ep/k1/rasp1.xlsx", tempPath, "temp_schedule.xlsx");

        string newFileHash = ComputeFileHash(Path.Combine(tempPath, "temp_schedule.xlsx"));

        // Проверяем содержимое ключевых ячеек
        bool isContentChanged = IsScheduleContentChanged(Path.Combine(tempPath, "temp_schedule.xlsx"), currentFilePath);

        if (newFileHash == previousFileHash && !isContentChanged)
        {
            Console.WriteLine("Расписание не изменилось.");
            System.IO.File.Delete(Path.Combine(tempPath, "temp_schedule.xlsx"));
            return;
        }

        Console.WriteLine("Обновлено расписание, рассылаем уведомления подписчикам...");
        previousFileHash = newFileHash;
        currentFilePath = await DownloadAndSaveFileAsync("https://serp-koll.ru/images/ep/k1/rasp1.xlsx", directoryPath);


        foreach (var sub in subscriptionManager.GetAllSubscriptions())
        {
            var chatId = sub.Key;
            var (type, name) = sub.Value;
            string message;
            var format = formatManager.GetFormat(chatId); // Получаем формат расписания (текст или фото)

            try
            {
                if (type == "group")
                {
                    message = GetGroupSchedule(currentFilePath, name);
                }
                else if (type == "teacher")
                {
                    message = GetTeacherSchedule(currentFilePath, name);
                }
                else
                {
                    message = "Неверный тип подписки.";
                }

                if (format == "photo")
                {
                    // Генерация изображения расписания
                    var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{name}_schedule.png");
                    Converter.ConvertScheduleToImage(name, message, imagePath);

                    // Отправка изображения
                    await using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                    await botClient.SendPhotoAsync(
                        chatId,
                        photo: stream,
                        caption: $"Обновлённое расписание для {name}",
                        cancellationToken: CancellationToken.None
                    );
                }
                else
                {
                    // Отправка текстового расписания
                    await botClient.SendTextMessageAsync(
                        chatId,
                        $"Обновлённое расписание:\n{message}",
                        cancellationToken: CancellationToken.None
                    );
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при отправке расписания для {name}: {ex.Message}");
                await botClient.SendTextMessageAsync(chatId, "Произошла ошибка при обновлении расписания.", cancellationToken: CancellationToken.None);
            }
        }

    }
    private static bool IsScheduleContentChanged(string newFilePath, string oldFilePath)
    {
        try
        {
            using var newPackage = new ExcelPackage(new FileInfo(newFilePath));
            using var oldPackage = new ExcelPackage(new FileInfo(oldFilePath));

            var newWorksheet = newPackage.Workbook.Worksheets.FirstOrDefault();
            var oldWorksheet = oldPackage.Workbook.Worksheets.FirstOrDefault();

            if (newWorksheet == null || oldWorksheet == null)
            {
                Console.WriteLine("Один из файлов пуст, считаем расписание обновлённым.");
                return true;
            }

            // Проверяем ключевые ячейки, например, дату расписания
            string newDate = newWorksheet.Cells[1, 2].Text.Trim();
            string oldDate = oldWorksheet.Cells[1, 2].Text.Trim();

            if (!string.Equals(newDate, oldDate, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Дата расписания изменилась: {oldDate} -> {newDate}");
                return true;
            }

            // Проверяем заголовки групп (строка 2)
            for (int col = 2; col <= newWorksheet.Dimension.End.Column; col++)
            {
                string newGroup = newWorksheet.Cells[2, col].Text.Trim();
                string oldGroup = oldWorksheet.Cells[2, col].Text.Trim();

                if (!string.Equals(newGroup, oldGroup, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Изменился список групп: {oldGroup} -> {newGroup}");
                    return true;
                }
            }

            // Проверяем первые 5 строк с расписанием
            for (int row = 3; row <= Math.Min(7, newWorksheet.Dimension.End.Row); row++)
            {
                for (int col = 2; col <= newWorksheet.Dimension.End.Column; col++)
                {
                    string newCell = newWorksheet.Cells[row, col].Text.Trim();
                    string oldCell = oldWorksheet.Cells[row, col].Text.Trim();

                    if (!string.Equals(newCell, oldCell, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Изменение в расписании: [{row}, {col}] {oldCell} -> {newCell}");
                        return true;
                    }
                }
            }

            return false; // Содержимое не изменилось
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при сравнении содержимого файлов: {ex.Message}");
            return true; // Если произошла ошибка, считаем, что содержимое изменилось
        }
    }




    private static string ComputeFileHash(string filePath)
    {
        using var md5 = MD5.Create();
        using var stream = System.IO.File.OpenRead(filePath);
        var hash = md5.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }

    private static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
    {
        // Обработка сообщений
        if (update.Type == UpdateType.Message && update.Message?.Text != null)
        {
            var chatId = update.Message.Chat.Id;
            var messageText = update.Message.Text.Trim();
            var previousSchedulePath = GetPreviousSchedulePath(chatId);

            switch (messageText)
            {
                case "/start":
                    // Убедимся, что рассылка подключена только через /start, иначе ничего не делаем.
                    await botClient.SendTextMessageAsync(chatId,
                        "Хотите получать рассылку?\nДля группы или для преподавателя?",
                        replyMarkup: new InlineKeyboardMarkup(new[]
                        {
                            new[] { InlineKeyboardButton.WithCallbackData("Для группы", "choose_group") },
                            new[] { InlineKeyboardButton.WithCallbackData("Для преподавателя", "choose_teacher") },
                            new[] { InlineKeyboardButton.WithCallbackData("Пропустить рассылку", "skip_subscription") }
                        }),
                        cancellationToken: cancellationToken);
                    break;


                case "choose_group":
                    await SendGroupList(chatId, cancellationToken, currentFilePath);
                    break;

                case "choose_teacher":
                    await SendTeacherList(chatId, cancellationToken, currentFilePath);
                    break;

                case "/group":
                    await SendGroupList(chatId, cancellationToken, currentFilePath);
                    Console.WriteLine("Проверка обновления расписания перед отправкой списка групп...");
                    await UpdateScheduleAndNotify();  // Обновляем расписание и уведомляем
                    break;

                case "/teacher":
                    await SendTeacherList(chatId, cancellationToken, currentFilePath);
                    Console.WriteLine("Проверка обновления расписания перед отправкой списка преподавателей...");
                    await UpdateScheduleAndNotify();  // Обновляем расписание и уведомляем
                    break;

                case "/full":
                    await botClient.SendDocumentAsync(chatId: chatId,
                        document: new InputFileStream(new MemoryStream(await DownloadFile()), "Расписание.xlsx"),
                        caption: "Вот файл с расписанием.",
                        cancellationToken: cancellationToken);
                    break;
                case "/settings":
                    await botClient.SendTextMessageAsync(chatId,
                        "Настройки:",
                        replyMarkup: new InlineKeyboardMarkup(new[]
                        {
                        new[] { InlineKeyboardButton.WithCallbackData("Расписание в виде ФОТО", "photo_schedule") },
                        new[] { InlineKeyboardButton.WithCallbackData("Расписание в виде ТЕКСТА", "text_schedule") },
                        new[] { InlineKeyboardButton.WithCallbackData("Отключить рассылку", "disable_subscription") }
                        }),
                        cancellationToken: cancellationToken);
                    break;
                case "/return":
                    // Запрашиваем у пользователя, чьё старое расписание он хочет получить
                    await botClient.SendTextMessageAsync(chatId,
                        "Чьё расписание вы хотите посмотреть?",
                        replyMarkup: new InlineKeyboardMarkup(new[]
                        {
                            new[] { InlineKeyboardButton.WithCallbackData("Группы", "group_old") },
                            new[] { InlineKeyboardButton.WithCallbackData("Преподавателя", "teacher_old") }
                        }),
                        cancellationToken: cancellationToken);
                    break;
                case "/group_old":
                    await SendGroupListOld(chatId, cancellationToken, previousSchedulePath);
                    break;
                case "/teacher_old":
                    await SendTeacherListOld(chatId, cancellationToken, previousSchedulePath);
                    break;





                default:
                    await botClient.SendTextMessageAsync(chatId, "Неизвестная команда. Пожалуйста, выберите из меню.", cancellationToken: cancellationToken);
                    break;
            }
        }

        // Обработка callback (инлайн-кнопок)
        if (update.Type == UpdateType.CallbackQuery)
        {
            var callbackQuery = update.CallbackQuery;
            var chatId = callbackQuery.Message.Chat.Id;

            if (callbackQuery.Data == "choose_group")
            {
                subscriptionManager.SetState(chatId, "choose_group");
                await SendGroupList(chatId, cancellationToken, currentFilePath); // Отправляем список групп
            }
            else if (callbackQuery.Data == "choose_teacher")
            {
                subscriptionManager.SetState(chatId, "choose_teacher");
                await SendTeacherList(chatId, cancellationToken, currentFilePath); // Отправляем список преподавателей
            }
            else if (callbackQuery.Data == "skip_subscription")
            {
                await botClient.SendTextMessageAsync(chatId, "Вы выбрали пропустить рассылку. Вы всегда можете вернуться к настройке рассылки через /start.", cancellationToken: cancellationToken);
                subscriptionManager.RemoveSubscription(chatId); // Убираем рассылку
                formatManager.RemoveFormat(chatId); // Убираем формат
            }
            else if (callbackQuery.Data == "photo_schedule")
            {
                formatManager.SetFormat(chatId, "photo");
                await botClient.SendTextMessageAsync(chatId, "Формат изменён на фото.", cancellationToken: cancellationToken);
            }
            else if (callbackQuery.Data == "text_schedule")
            {
                formatManager.SetFormat(chatId, "text");
                await botClient.SendTextMessageAsync(chatId, "Формат изменён на текст.", cancellationToken: cancellationToken);
            }
            else if (callbackQuery.Data == "disable_subscription")
            {
                subscriptionManager.RemoveSubscription(chatId);
                formatManager.RemoveFormat(chatId);
                await botClient.SendTextMessageAsync(chatId, "Рассылка отключена.", cancellationToken: cancellationToken);
            }
            else if (callbackQuery.Data.StartsWith("group_old_"))
            {
                var groupName = callbackQuery.Data.Replace("group_old_", "");

                // Получаем путь к старому расписанию для этой группы
                var previousSchedulePath = GetPreviousSchedulePath(chatId);
                if (previousSchedulePath != null)
                {
                    var schedule = GetGroupSchedule(previousSchedulePath, groupName); // Получаем старое расписание
                    var format = formatManager.GetFormat(chatId);

                    if (string.IsNullOrEmpty(schedule))
                    {
                        await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание для этой группы не найдено.", cancellationToken: cancellationToken);
                        return;
                    }

                    // Отправка старого расписания в нужном формате (текст или фото)
                    if (format == "photo")
                    {
                        var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{groupName}_schedule.png");
                        Converter.ConvertScheduleToImage(groupName, schedule, imagePath);

                        await using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                        await botClient.SendPhotoAsync(
                            chatId,
                            photo: stream,
                            caption: $"Предыдущее расписание для группы {groupName}",
                            cancellationToken: cancellationToken
                        );
                    }
                    else
                    {
                        await botClient.SendTextMessageAsync(chatId, schedule, cancellationToken: cancellationToken);
                    }
                }
                else
                {
                    await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание для группы не найдено.", cancellationToken: cancellationToken);
                }
            }
            else if (callbackQuery.Data.StartsWith("teacher_old_"))
            {
                var teacherName = callbackQuery.Data.Replace("teacher_old_", "");

                // Получаем путь к старому расписанию для этого преподавателя
                var previousSchedulePath = GetPreviousSchedulePath(chatId);
                if (previousSchedulePath != null)
                {
                    var schedule = GetTeacherSchedule(previousSchedulePath, teacherName); // Получаем старое расписание
                    var format = formatManager.GetFormat(chatId);

                    if (string.IsNullOrEmpty(schedule))
                    {
                        await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание для этого преподавателя не найдено.", cancellationToken: cancellationToken);
                        return;
                    }

                    // Отправка старого расписания в нужном формате (текст или фото)
                    if (format == "photo")
                    {
                        var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{teacherName}_schedule.png");
                        Converter.ConvertScheduleToImage(teacherName, schedule, imagePath);

                        await using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                        await botClient.SendPhotoAsync(
                            chatId,
                            photo: stream,
                            caption: $"Предыдущее расписание для преподавателя {teacherName}",
                            cancellationToken: cancellationToken
                        );
                    }
                    else
                    {
                        await botClient.SendTextMessageAsync(chatId, schedule, cancellationToken: cancellationToken);
                    }
                }
                else
                {
                    await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание для преподавателя не найдено.", cancellationToken: cancellationToken);
                }
            }
            else if (callbackQuery.Data.StartsWith("group_"))
            {
                var groupName = callbackQuery.Data.Replace("group_", "");
                var currentFileName = Path.GetFileName(currentFilePath); // например: "Schedule_13_12_2024.xlsx"

                // Извлекаем дату из имени файла: начиная с 9-го символа (после "Schedule_")
                var currentDateString = currentFileName.Substring(9, 10).Trim(); // "13_12_2024"

                // Преобразуем строку в DateTime
                DateTime currentFileDate = DateTime.ParseExact(currentDateString, "dd_MM_yyyy", null, System.Globalization.DateTimeStyles.None);



                // Получаем текущее расписание для группы
                var schedule = GetGroupSchedule(currentFilePath, groupName);
                var format = formatManager.GetFormat(chatId);

                if (string.IsNullOrEmpty(schedule))
                {
                    await botClient.SendTextMessageAsync(chatId, "Расписание для этой группы не найдено.", cancellationToken: cancellationToken);
                    return;
                }

                // Отправка расписания в нужном формате (текст или фото)
                if (format == "photo")
                {
                    var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{groupName}_schedule.png");
                    Converter.ConvertScheduleToImage(groupName, schedule, imagePath);

                    await using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                    await botClient.SendPhotoAsync(
                        chatId,
                        photo: stream,
                        caption: $"Текущее расписание для группы {groupName} на {currentFileDate.ToString("dd.MM.yyyy")}",
                        cancellationToken: cancellationToken
                    );
                }
                else
                {
                    await botClient.SendTextMessageAsync(chatId, schedule, cancellationToken: cancellationToken);
                }
            }
            else if (callbackQuery.Data.StartsWith("teacher_"))
            {
                var teacherName = callbackQuery.Data.Replace("teacher_", "");
                var currentFileName = Path.GetFileName(currentFilePath); // например: "Schedule_13_12_2024.xlsx"
                var currentDateString = currentFileName.Substring(9, 10).Trim(); // "13_12_2024"
                DateTime currentFileDate = DateTime.ParseExact(currentDateString, "dd_MM_yyyy", null, System.Globalization.DateTimeStyles.None);

                // Получаем текущее расписание для преподавателя
                var schedule = GetTeacherSchedule(currentFilePath, teacherName);
                var format = formatManager.GetFormat(chatId);

                if (string.IsNullOrEmpty(schedule))
                {
                    await botClient.SendTextMessageAsync(chatId, "Расписание для этого преподавателя не найдено.", cancellationToken: cancellationToken);
                    return;
                }

                // Отправка расписания в нужном формате (текст или фото)
                if (format == "photo")
                {
                    var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{teacherName}_schedule.png");
                    Converter.ConvertScheduleToImage(teacherName, schedule, imagePath);

                    await using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                    await botClient.SendPhotoAsync(
                        chatId,
                        photo: stream,
                        caption: $"Текущее расписание для преподавателя {teacherName} на {currentFileDate.ToString("dd.MM.yyyy")}",
                        cancellationToken: cancellationToken
                    );
                }
                else
                {
                    await botClient.SendTextMessageAsync(chatId, schedule, cancellationToken: cancellationToken);
                }
            }
            








        }
    }
    



    private static void RemoveSubscription(long chatId)
    {
        // Логика удаления подписки из текстового файла или базы данных
        var subscriptionFilePath = "subscriptions.txt";

        if (System.IO.File.Exists(subscriptionFilePath))
        {
            var lines = System.IO.File.ReadAllLines(subscriptionFilePath).Where(line => !line.StartsWith(chatId.ToString())).ToList();
            System.IO.File.WriteAllLines(subscriptionFilePath, lines);
        }
    }


    private static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
    {
        Console.WriteLine($"Ошибка: {exception.Message}");
        return Task.CompletedTask;
    }

    private static async Task SendGroupList(long chatId, CancellationToken cancellationToken, string filePath)
    {
        var groups = GetGroups(filePath);
        var keyboardButtons = groups
            .Select(group => InlineKeyboardButton.WithCallbackData(group, $"group_{group}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        var keyboard = new InlineKeyboardMarkup(keyboardButtons);
        await botClient.SendTextMessageAsync(chatId, "Выберите группу:", replyMarkup: keyboard, cancellationToken: cancellationToken);
    }
    private static async Task SendGroupListOld(long chatId, CancellationToken cancellationToken, string filePath)
    {
        var groups = GetGroups(filePath);
        var keyboardButtons = groups
            .Select(group => InlineKeyboardButton.WithCallbackData(group, $"group_old_{group}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        var keyboard = new InlineKeyboardMarkup(keyboardButtons);
        await botClient.SendTextMessageAsync(chatId, "Выберите группу:", replyMarkup: keyboard, cancellationToken: cancellationToken);
    }

    private static async Task SendTeacherList(long chatId, CancellationToken cancellationToken, string filePath)
    {
        var teachers = GetTeachers(filePath);
        var keyboardButtons = teachers
            .Select(teacher => InlineKeyboardButton.WithCallbackData(teacher, $"teacher_{teacher}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        var keyboard = new InlineKeyboardMarkup(keyboardButtons);

        await botClient.SendTextMessageAsync(chatId, "Выберите преподавателя:", replyMarkup: keyboard, cancellationToken: cancellationToken);
    }
    private static async Task SendTeacherListOld(long chatId, CancellationToken cancellationToken, string filePath)
    {
        var teachers = GetTeachers(filePath);
        var keyboardButtons = teachers
            .Select(teacher => InlineKeyboardButton.WithCallbackData(teacher, $"teacher_old_{teacher}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        var keyboard = new InlineKeyboardMarkup(keyboardButtons);

        await botClient.SendTextMessageAsync(chatId, "Выберите преподавателя:", replyMarkup: keyboard, cancellationToken: cancellationToken);
    }

    private static List<string> GetGroups(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        var groups = new List<string>();
        for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
        {
            var groupName = worksheet.Cells[2, col].Text;
            if (!string.IsNullOrEmpty(groupName))
            {
                groups.Add(groupName);
            }
        }
        return groups;
    }

    private static List<string> GetTeachers(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        var teachers = new HashSet<string>();

        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
            {
                var cellValue = worksheet.Cells[row, col].Text;
                if (string.IsNullOrWhiteSpace(cellValue))
                    continue;

                var parts = cellValue.Split('\n')
                                     .Select(p => p.Trim())
                                     .Where(p => !string.IsNullOrEmpty(p))
                                     .ToArray();

                if (parts.Length >= 3)
                {
                    var teacher = parts[2];
                    if (!string.IsNullOrEmpty(teacher))
                    {
                        teachers.Add(teacher);
                    }
                }
            }
        }

        return teachers.ToList();
    }

    private static string GetGroupSchedule(string filePath, string groupName)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        string dateRaw = worksheet.Cells[1, 2].Text.Trim();
        string prefix = "Расписание на ";
        string dateOnly = dateRaw.StartsWith(prefix) ? dateRaw.Substring(prefix.Length) : dateRaw;
        string dayOfWeek = worksheet.Cells[1, 16].Text.Trim();

        int groupColumn = -1;
        for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
        {
            if (worksheet.Cells[2, col].Text == groupName)
            {
                groupColumn = col;
                break;
            }
        }

        if (groupColumn == -1)
            return $"Расписание для группы {groupName} не найдено.";

        var schedule = new List<string>();

        string header = $"Расписание для группы {groupName} на {dateOnly}";
        if (!string.IsNullOrEmpty(dayOfWeek))
        {
            header += $"{dayOfWeek}";
        }
        header += ":";

        schedule.Add(header);

        var pairRows = new[]
        {
            new[] {14, 15},
            new[] {27, 28},
            new[] {40, 41},
            new[] {53, 54},
            new[] {66, 67}
        };

        string ParseCellValue(string cellValue)
        {
            cellValue = cellValue.Trim();
            if (string.IsNullOrEmpty(cellValue))
                return null;

            var parts = cellValue.Split('\n').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray();

            string code = parts.Length > 0 ? parts[0] : "";
            string subjectName = parts.Length > 1 ? parts[1] : "";
            string teacher = parts.Length > 2 ? parts[2] : "";
            string room = parts.Length > 3 ? parts[3].Trim('(', ')') : "";

            if (string.IsNullOrEmpty(code) && string.IsNullOrEmpty(subjectName) && string.IsNullOrEmpty(teacher) && string.IsNullOrEmpty(room))
                return null;

            var lines = new List<string>();
            if (!string.IsNullOrEmpty(code) || !string.IsNullOrEmpty(subjectName))
            {
                string firstLine = code;
                if (!string.IsNullOrEmpty(subjectName))
                    firstLine += (string.IsNullOrEmpty(firstLine) ? "" : " ") + subjectName;
                lines.Add(firstLine);
            }

            lines.Add(string.IsNullOrEmpty(teacher) ? "Преподаватель: -" : $"Преподаватель: {teacher}");
            lines.Add(string.IsNullOrEmpty(room) ? "Кабинет: -" : $"Кабинет: {room}");

            return string.Join("\n", lines);
        }

        for (int i = 0; i < pairRows.Length; i++)
        {
            int pairNumber = i + 1;
            int rowA = pairRows[i][0];
            int rowB = pairRows[i][1];

            var cellA = worksheet.Cells[rowA, groupColumn].Text;
            var cellB = worksheet.Cells[rowB, groupColumn].Text;
            var cellAMerge = worksheet.Cells[rowA, groupColumn];

            string parsedA = ParseCellValue(cellA);
            string parsedB = ParseCellValue(cellB);

            if (parsedA == null && parsedB == null)
            {
                schedule.Add($"{pairNumber}) Нет пары");
            }
            else if (cellAMerge.Merge)
            {
                schedule.Add($"{pairNumber}) {parsedA}");
            }
            else if (parsedA != null && parsedB == null && !cellAMerge.Merge)
            {
                var linesA = parsedA.Split('\n');

                schedule.Add($"{pairNumber}) Разделённая пара:");
                schedule.Add($"Подгруппа 1: {linesA[0]}");
                for (int la = 1; la < linesA.Length; la++)
                {
                    schedule.Add(linesA[la]);
                }

                schedule.Add("Подгруппа 2: Нет пары");
      
            }
            else if (parsedA == null && parsedB != null)
            {
                var linesB = parsedB.Split('\n');

                schedule.Add($"{pairNumber}) Разделённая пара:");
                schedule.Add("Подгруппа 1: Нет пары");
                schedule.Add($"Подгруппа 2: {linesB[0]}");
                for (int lb = 1; lb < linesB.Length; lb++)
                {
                    schedule.Add(linesB[lb]);
                }
            }
            else
            {
                var linesA = parsedA.Split('\n');
                var linesB = parsedB.Split('\n');

                schedule.Add($"{pairNumber}) Разделённая пара:");
                schedule.Add($"Подгруппа 1: {linesA[0]}");
                for (int la = 1; la < linesA.Length; la++)
                {
                    schedule.Add(linesA[la]);
                }

                schedule.Add($"Подгруппа 2: {linesB[0]}");
                for (int lb = 1; lb < linesB.Length; lb++)
                {
                    schedule.Add(linesB[lb]);
                }
            }

            schedule.Add("");
        }

        return string.Join("\n", schedule);
    }

    private static string GetTeacherSchedule(string filePath, string teacherName)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        string dateRaw = worksheet.Cells[1, 2].Text.Trim();
        string prefix = "Расписание на ";
        string dateOnly = dateRaw.StartsWith(prefix) ? dateRaw.Substring(prefix.Length) : dateRaw;

        string dayOfWeek = worksheet.Cells[1, 16].Text.Trim();

        var schedule = new List<string>();
        string header = $"Расписание для преподавателя {teacherName} на {dateOnly}";
        if (!string.IsNullOrEmpty(dayOfWeek))
        {
            header += $"{dayOfWeek}";
        }
        header += ":";
        schedule.Add(header);

        var pairRows = new[]
        {
            new[] {14, 15},
            new[] {27, 28},
            new[] {40, 41},
            new[] {53, 54},
            new[] {66, 67}
        };

        (string code, string subject, string teacher, string room)? ParseCell(string cellValue)
        {
            if (string.IsNullOrWhiteSpace(cellValue))
                return null;

            var parts = cellValue.Split('\n').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray();
            if (parts.Length == 0) return null;

            string code = parts.Length > 0 ? parts[0] : "";
            string subj = parts.Length > 1 ? parts[1] : "";
            string tchr = parts.Length > 2 ? parts[2] : "";
            string rm = parts.Length > 3 ? parts[3].Trim('(', ')') : "";

            if (string.IsNullOrEmpty(code) && string.IsNullOrEmpty(subj) && string.IsNullOrEmpty(tchr) && string.IsNullOrEmpty(rm))
                return null;

            return (code, subj, tchr, rm);
        }

        int startCol = 2;
        int endCol = worksheet.Dimension.End.Column;

        for (int i = 0; i < pairRows.Length; i++)
        {
            int pairNumber = i + 1;
            int rowA = pairRows[i][0];
            int rowB = pairRows[i][1];

            bool pairFound = false;

            for (int col = startCol; col <= endCol; col++)
            {
                string groupName = worksheet.Cells[2, col].Text.Trim();
                if (string.IsNullOrEmpty(groupName)) continue;

                var cellA = worksheet.Cells[rowA, col].Text;
                var cellB = worksheet.Cells[rowB, col].Text;

                var parsedA = ParseCell(cellA);
                var parsedB = ParseCell(cellB);

                void AddSubgroupInfo((string code, string subject, string teacher, string room) data)
                {
                    if (!pairFound)
                    {
                        schedule.Add($"{pairNumber}) Пара:");
                        pairFound = true;
                    }

                    var subjLine = $"{data.code} {data.subject}".Trim();
                    if (string.IsNullOrEmpty(subjLine))
                        subjLine = "-";

                    schedule.Add(subjLine);
                    schedule.Add($"Группа: {groupName}");
                    schedule.Add($"Преподаватель: {data.teacher}");
                    schedule.Add($"Кабинет: {(string.IsNullOrEmpty(data.room) ? "-" : data.room)}");
                    schedule.Add("");
                }

                if (parsedA.HasValue && parsedA.Value.teacher == teacherName)
                {
                    AddSubgroupInfo(parsedA.Value);
                }

                if (parsedB.HasValue && parsedB.Value.teacher == teacherName)
                {
                    AddSubgroupInfo(parsedB.Value);
                }
            }
        }

        return string.Join("\n", schedule);
    }

    private static async Task<byte[]> DownloadFile()
    {
        using var httpClient = new HttpClient();
        var response = await httpClient.GetAsync("https://serp-koll.ru/images/ep/k1/rasp1.xlsx");
        return await response.Content.ReadAsByteArrayAsync();
    }

    private static async Task<string> DownloadAndSaveFileAsync(string url, string directory, string fileName = null)
    {
        try
        {
            using var httpClient = new HttpClient();
            var response = await httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            // Скачиваем файл как массив байтов
            var fileBytes = await response.Content.ReadAsByteArrayAsync();

            // Создаём временный путь для сохранения файла
            var tempPath = Path.Combine(directory, "temp.xlsx");
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            await System.IO.File.WriteAllBytesAsync(tempPath, fileBytes);

            // Если имя файла не передано, читаем дату из файла и формируем имя
            if (string.IsNullOrEmpty(fileName))
            {
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(tempPath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet != null)
                    {
                        var dateCell = worksheet.Cells["B1"].Text.Trim();
                        if (dateCell.StartsWith("Расписание на "))
                        {
                            var extractedDate = dateCell.Replace("Расписание на ", "")
                                                        .Replace(" г.", "") // Убираем "г."
                                                        .Split('(', StringSplitOptions.RemoveEmptyEntries)[0]
                                                        .Trim();

                            // Преобразуем дату в формат для имени файла
                            var dateParts = extractedDate.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                            if (dateParts.Length == 3)
                            {
                                var day = dateParts[0];
                                var month = GetMonthNumber(dateParts[1]);
                                var year = dateParts[2];
                                if (month == null)
                                {
                                    throw new Exception("Не удалось преобразовать название месяца.");
                                }

                                fileName = $"Schedule_{day}_{month}_{year}.xlsx";
                            }
                            else
                            {
                                throw new Exception("Дата в ячейке B1 имеет неверный формат.");
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(fileName))
                {
                    throw new Exception("Не удалось извлечь дату для имени файла.");
                }
            }

            // Полный путь для сохранения файла
            var fullPath = Path.Combine(directory, fileName);

            // Перемещаем временный файл в нужное место с именем
            if (System.IO.File.Exists(fullPath))
            {
                System.IO.File.Delete(fullPath);
            }
            System.IO.File.Move(tempPath, fullPath);

            Console.WriteLine($"Файл успешно сохранён: {fullPath}");
            return fullPath;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при загрузке файла: {ex.Message}");
            throw;
        }
    }

    // Метод для преобразования названия месяца в номер
    private static string GetMonthNumber(string monthName)
    {
        return monthName.ToLower() switch
        {
            "января" => "01",
            "февраля" => "02",
            "марта" => "03",
            "апреля" => "04",
            "мая" => "05",
            "июня" => "06",
            "июля" => "07",
            "августа" => "08",
            "сентября" => "09",
            "октября" => "10",
            "ноября" => "11",
            "декабря" => "12",
            _ => null
        };
    }
    private static string GetPreviousSchedulePath(long chatId)
    {
        // Получаем имя текущего файла из currentFilePath
        var currentFileName = Path.GetFileName(currentFilePath); // например: "Schedule_13_12_2024.xlsx"

        // Извлекаем дату из имени файла: начиная с 9-го символа (после "Schedule_")
        var currentDateString = currentFileName.Substring(9, 10).Trim(); // "13_12_2024"

        // Преобразуем строку в DateTime
        DateTime currentFileDate;
        if (!DateTime.TryParseExact(currentDateString, "dd_MM_yyyy", null, System.Globalization.DateTimeStyles.None, out currentFileDate))
        {
            return null; // Если дата в имени файла невалидна, возвращаем null
        }

        // Получаем список всех файлов в папке, которые соответствуют формату расписания
        var scheduleFiles = Directory.GetFiles(directoryPath, "Schedule_*.xlsx");

        if (scheduleFiles.Length == 0)
        {
            return null; // Если нет файлов, возвращаем null
        }

        // Фильтруем файлы, оставляя только те, чьи даты раньше текущей
        var validFiles = scheduleFiles
            .Where(file => DateTime.TryParseExact(Path.GetFileName(file).Substring(9, 10), "dd_MM_yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime fileDate)
                          && fileDate < currentFileDate) // Только те файлы, чьи даты меньше текущей
            .ToList();

        if (validFiles.Count == 0)
        {
            return null; // Если нет файлов с датой раньше текущей, возвращаем null
        }

        // Сортировка по дате (по убыванию), чтобы выбрать последний файл с датой раньше текущей
        var latestFile = validFiles
            .OrderByDescending(file => DateTime.ParseExact(Path.GetFileName(file).Substring(9, 10), "dd_MM_yyyy", null)) // Сортируем по дате
            .First(); // Берем последний файл

        return latestFile; // Возвращаем путь к найденному файлу
    }

    private static async Task SendScheduleAsText(long chatId, string filePath, CancellationToken cancellationToken)
    {
        var schedule = await System.IO.File.ReadAllTextAsync(filePath, cancellationToken);
        await botClient.SendTextMessageAsync(chatId, schedule, cancellationToken: cancellationToken);
    }

    private static async Task SendScheduleAsImage(long chatId, string filePath, CancellationToken cancellationToken)
    {
        var groupName = "Группа"; // Здесь нужно определить, какую группу использовать (например, если расписание относится к группе)

        // Конвертируем текст расписания в изображение
        string outputImagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_previous_schedule.png");
        var scheduleText = await System.IO.File.ReadAllTextAsync(filePath, cancellationToken);

        // Конвертация расписания в изображение
        Converter.ConvertScheduleToImage(groupName, scheduleText, outputImagePath);

        // Отправка изображения
        await using var stream = new FileStream(outputImagePath, FileMode.Open, FileAccess.Read);
        await botClient.SendPhotoAsync(chatId, photo: stream, caption: "Предыдущее расписание", cancellationToken: cancellationToken);
    }







}
