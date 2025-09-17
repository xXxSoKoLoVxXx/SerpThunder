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
using System.Text;

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

        currentFilePath = await DownloadAndSaveFileAsync("https://serp-koll.ru/images/ep/k2/rasp2.xlsx", directoryPath);
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
            cts.Cancel();
        };

        try
        {
            await Task.Delay(Timeout.Infinite, cts.Token);
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

        await DownloadAndSaveFileAsync("https://serp-koll.ru/images/ep/k2/rasp2.xlsx", tempPath, "temp_schedule.xlsx");

        string newFileHash = ComputeFileHash(Path.Combine(tempPath, "temp_schedule.xlsx"));

        if (newFileHash == previousFileHash)
        {
            Console.WriteLine("Расписание не изменилось.");
            System.IO.File.Delete(Path.Combine(tempPath, "temp_schedule.xlsx"));
            return;
        }

        Console.WriteLine("Обновлено расписание, рассылаем уведомления подписчикам...");
        previousFileHash = newFileHash;
        currentFilePath = await DownloadAndSaveFileAsync("https://serp-koll.ru/images/ep/k2/rasp2.xlsx", directoryPath);

        foreach (var sub in subscriptionManager.GetAllSubscriptions())
        {
            var chatId = sub.Key;
            var (type, name) = sub.Value;
            string message;
            var format = formatManager.GetFormat(chatId);

            try
            {
                if (type == "group")
                {
                    message = GetGroupSchedule(currentFilePath, name);
                }
                else
                {
                    continue; // Пропускаем неподдерживаемые типы
                }

                if (format == "photo")
                {
                    var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{name}_schedule.png");
                    Converter.ConvertScheduleToImage(name, message, imagePath, currentFilePath);

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

    private static string ComputeFileHash(string filePath)
    {
        using var md5 = MD5.Create();
        using var stream = System.IO.File.OpenRead(filePath);
        var hash = md5.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }

    private static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
    {
        if (update.Type == UpdateType.Message && update.Message?.Text != null)
        {
            var chatId = update.Message.Chat.Id;
            var messageText = update.Message.Text.Trim();

            switch (messageText)
            {
                case "/start":
                    await botClient.SendTextMessageAsync(chatId,
                        "Добро пожаловать! Выберите группу для подписки на рассылку:",
                        replyMarkup: new InlineKeyboardMarkup(new[]
                        {
                            InlineKeyboardButton.WithCallbackData("Выбрать группу", "start_subscribe")
                        }),
                        cancellationToken: cancellationToken);
                    break;

                case "/group":
                    await UpdateScheduleAndNotify(); // Обновляем расписание перед показом
                    await botClient.SendTextMessageAsync(chatId,
                        "Выберите группу:",
                        replyMarkup: await CreateGroupKeyboard(currentFilePath),
                        cancellationToken: cancellationToken);
                    break;

                case "/group_old":
                    var previousSchedulePath = GetPreviousSchedulePath(chatId);
                    if (previousSchedulePath != null)
                    {
                        await botClient.SendTextMessageAsync(chatId,
                            "Выберите группу для просмотра предыдущего расписания:",
                            replyMarkup: await CreateGroupKeyboardOld(previousSchedulePath),
                            cancellationToken: cancellationToken);
                    }
                    else
                    {
                        await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание не найдено.", cancellationToken: cancellationToken);
                    }
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

                default:
                    await botClient.SendTextMessageAsync(chatId, 
                        "Доступные команды:\n/start - Подписаться на рассылку\n/group - Посмотреть расписание группы\n/group_old - Предыдущее расписание\n/full - Скачать файл расписания\n/settings - Настройки", 
                        cancellationToken: cancellationToken);
                    break;
            }
        }

        if (update.Type == UpdateType.CallbackQuery)
        {
            var callbackQuery = update.CallbackQuery;
            var chatId = callbackQuery.Message.Chat.Id;

            if (callbackQuery.Data.StartsWith("subscribe_"))
            {
                var groupName = callbackQuery.Data.Replace("subscribe_", "");
                subscriptionManager.AddSubscription(chatId, "group", groupName);
                await botClient.SendTextMessageAsync(chatId, $"Вы успешно подписаны на рассылку для группы {groupName}.", cancellationToken: cancellationToken);
            }
            else if (callbackQuery.Data == "start_subscribe")
            {
                await botClient.SendTextMessageAsync(chatId,
                    "Выберите группу для подписки:",
                    replyMarkup: await CreateSubscribeKeyboard(currentFilePath),
                    cancellationToken: cancellationToken);
            }
            else if (callbackQuery.Data.StartsWith("group_"))
            {
                var groupName = callbackQuery.Data.Replace("group_", "");
                var schedule = GetGroupSchedule(currentFilePath, groupName);
                var format = formatManager.GetFormat(chatId);

                if (string.IsNullOrEmpty(schedule))
                {
                    await botClient.SendTextMessageAsync(chatId, "Расписание для этой группы не найдено.", cancellationToken: cancellationToken);
                    return;
                }

                if (format == "photo")
                {
                    var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{groupName}_schedule.png");
                    Converter.ConvertScheduleToImage(groupName, schedule, imagePath, currentFilePath);

                    await using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                    await botClient.SendPhotoAsync(
                        chatId,
                        photo: stream,
                        caption: $"Расписание для группы {groupName}",
                        cancellationToken: cancellationToken
                    );
                }
                else
                {
                    await botClient.SendTextMessageAsync(chatId, schedule, cancellationToken: cancellationToken);
                }
            }
            else if (callbackQuery.Data.StartsWith("group_old_"))
            {
                var groupName = callbackQuery.Data.Replace("group_old_", "");
                var previousSchedulePath = GetPreviousSchedulePath(chatId);
                if (previousSchedulePath != null)
                {
                    var schedule = GetGroupSchedule(previousSchedulePath, groupName);
                    var format = formatManager.GetFormat(chatId);

                    if (string.IsNullOrEmpty(schedule))
                    {
                        await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание для этой группы не найдено.", cancellationToken: cancellationToken);
                        return;
                    }

                    if (format == "photo")
                    {
                        var imagePath = Path.Combine(Path.GetTempPath(), $"{chatId}_{groupName}_old_schedule.png");
                        Converter.ConvertScheduleToImage(groupName, schedule, imagePath, previousSchedulePath);

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
                    await botClient.SendTextMessageAsync(chatId, "Предыдущее расписание не найдено.", cancellationToken: cancellationToken);
                }
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
        }
    }

    private static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
    {
        Console.WriteLine($"Ошибка: {exception.Message}");
        return Task.CompletedTask;
    }

    private static async Task<InlineKeyboardMarkup> CreateSubscribeKeyboard(string filePath)
    {
        var groups = GetGroups(filePath);
        var keyboardButtons = groups
            .Select(group => InlineKeyboardButton.WithCallbackData(group, $"subscribe_{group}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        return new InlineKeyboardMarkup(keyboardButtons);
    }

    private static async Task<InlineKeyboardMarkup> CreateGroupKeyboard(string filePath)
    {
        var groups = GetGroups(filePath);
        var keyboardButtons = groups
            .Select(group => InlineKeyboardButton.WithCallbackData(group, $"group_{group}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        return new InlineKeyboardMarkup(keyboardButtons);
    }

    private static async Task<InlineKeyboardMarkup> CreateGroupKeyboardOld(string filePath)
    {
        var groups = GetGroups(filePath);
        var keyboardButtons = groups
            .Select(group => InlineKeyboardButton.WithCallbackData(group, $"group_old_{group}"))
            .Chunk(3)
            .Select(chunk => chunk.ToArray())
            .ToArray();

        return new InlineKeyboardMarkup(keyboardButtons);
    }

    private static List<string> GetGroups(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        var groups = new List<string>();
        
        // Проверяем разные варианты расположения групп
        // Вариант 1: C2:AA2 (понедельник с классным часом)
        for (int col = 3; col <= 27; col++) // C=3, AA=27
        {
            var groupName = worksheet.Cells[2, col].Text?.Trim();
            if (!string.IsNullOrEmpty(groupName) && !groups.Contains(groupName))
            {
                groups.Add(groupName);
            }
        }

        // Вариант 2 и 3: D2:AA2 (стандартные варианты)
        for (int col = 4; col <= 27; col++) // D=4, AA=27
        {
            var groupName = worksheet.Cells[2, col].Text?.Trim();
            if (!string.IsNullOrEmpty(groupName) && !groups.Contains(groupName))
            {
                groups.Add(groupName);
            }
        }

        // Вариант 3: второе расписание D9:AA9
        for (int col = 4; col <= 27; col++)
        {
            var groupName = worksheet.Cells[9, col].Text?.Trim();
            if (!string.IsNullOrEmpty(groupName) && !groups.Contains(groupName))
            {
                groups.Add(groupName);
            }
        }

        return groups.Where(g => !string.IsNullOrWhiteSpace(g)).Distinct().ToList();
    }

    private static string GetGroupSchedule(string filePath, string groupName)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        var schedules = new List<string>();

        // Ищем все возможные расписания
        var scheduleBlocks = DetectScheduleBlocks(worksheet);

        foreach (var block in scheduleBlocks)
        {
            var schedule = ParseScheduleBlock(worksheet, block, groupName);
            if (!string.IsNullOrEmpty(schedule))
            {
                schedules.Add(schedule);
            }
        }

        return string.Join("\n\n", schedules);
    }

    private static List<ScheduleBlock> DetectScheduleBlocks(ExcelWorksheet worksheet)
    {
        var blocks = new List<ScheduleBlock>();

        // Вариант 1: Понедельник с классным часом
        if (!string.IsNullOrEmpty(worksheet.Cells[3, 3].Text) && 
            worksheet.Cells[3, 3].Text.Contains("Классный час"))
        {
            blocks.Add(new ScheduleBlock
            {
                DateCell = "A1",
                GroupRow = 2,
                GroupStartCol = 3, // C
                GroupEndCol = 18, // R
                TimeStartRow = 3,
                TimeCol = 2, // B
                ScheduleStartRow = 3,
                ScheduleEndRow = 8,
                ScheduleStartCol = 3, // C
                ScheduleEndCol = 18, // R
                HasClassHour = true
            });
        }

        // Вариант 2: Стандартное расписание
        if (!string.IsNullOrEmpty(worksheet.Cells[1, 1].Text))
        {
            blocks.Add(new ScheduleBlock
            {
                DateCell = "A1",
                GroupRow = 2,
                GroupStartCol = 4, // D
                GroupEndCol = 27, // AA
                TimeStartRow = 3,
                TimeCol = 3, // C
                ScheduleStartRow = 3,
                ScheduleEndRow = 7,
                ScheduleStartCol = 4, // D
                ScheduleEndCol = 27, // AA
                HasClassHour = false
            });
        }

        // Вариант 3: Второе расписание внизу
        if (!string.IsNullOrEmpty(worksheet.Cells[8, 1].Text))
        {
            blocks.Add(new ScheduleBlock
            {
                DateCell = "A8",
                GroupRow = 9,
                GroupStartCol = 4, // D
                GroupEndCol = 27, // AA
                TimeStartRow = 10,
                TimeCol = 3, // C
                ScheduleStartRow = 10,
                ScheduleEndRow = 14,
                ScheduleStartCol = 4, // D
                ScheduleEndCol = 27, // AA
                HasClassHour = false
            });
        }

        return blocks;
    }

    private static string ParseScheduleBlock(ExcelWorksheet worksheet, ScheduleBlock block, string groupName)
    {
        // Находим колонку группы
        int groupColumn = -1;
        for (int col = block.GroupStartCol; col <= block.GroupEndCol; col++)
        {
            if (worksheet.Cells[block.GroupRow, col].Text?.Trim() == groupName)
            {
                groupColumn = col;
                break;
            }
        }

        if (groupColumn == -1)
            return "";

        var schedule = new List<string>();

        // Добавляем заголовок с датой
        string dateText = "";
        if (block.DateCell == "A1")
            dateText = worksheet.Cells[1, 1].Text?.Trim() ?? "";
        else if (block.DateCell == "A8")
            dateText = worksheet.Cells[8, 1].Text?.Trim() ?? "";

        schedule.Add($"Расписание для группы {groupName}");
        if (!string.IsNullOrEmpty(dateText))
            schedule.Add(dateText);
        schedule.Add("");

        // Парсим пары
        int pairNumber = block.HasClassHour ? 0 : 1;
        
        for (int row = block.ScheduleStartRow; row <= block.ScheduleEndRow; row++)
        {
            // Получаем время пары
            string time = worksheet.Cells[row, block.TimeCol].Text?.Trim() ?? "";
            
            // Получаем содержимое пары
            string cellContent = worksheet.Cells[row, groupColumn].Text?.Trim() ?? "";

            if (block.HasClassHour && row == block.ScheduleStartRow)
            {
                // Классный час
                schedule.Add("Классный час");
                if (!string.IsNullOrEmpty(time))
                    schedule.Add($"Время: {time}");
                
                // Кабинет для классного часа
                string classroom = worksheet.Cells[row + 1, groupColumn].Text?.Trim() ?? "";
                if (!string.IsNullOrEmpty(classroom))
                    schedule.Add($"Кабинет: {classroom}");
                
                schedule.Add("");
                row++; // Пропускаем строку с кабинетами
                continue;
            }

            if (!string.IsNullOrEmpty(cellContent) || !string.IsNullOrEmpty(time))
            {
                schedule.Add($"{pairNumber}) Пара:");
                
                if (!string.IsNullOrEmpty(time))
                    schedule.Add($"Время: {time}");

                if (!string.IsNullOrEmpty(cellContent))
                {
                    var parts = cellContent.Split('\n')
                                          .Select(p => p.Trim())
                                          .Where(p => !string.IsNullOrEmpty(p))
                                          .ToArray();

                    string subject = parts.Length > 0 ? parts[0] : "";
                    string teacher = parts.Length > 1 ? parts[1] : "";
                    string room = parts.Length > 2 ? parts[2] : "";

                    if (!string.IsNullOrEmpty(subject))
                        schedule.Add($"Предмет: {subject}");
                    if (!string.IsNullOrEmpty(teacher))
                        schedule.Add($"Преподаватель: {teacher}");
                    if (!string.IsNullOrEmpty(room))
                        schedule.Add($"Кабинет: {room}");
                }
                else
                {
                    schedule.Add("Нет пары");
                }

                schedule.Add("");
            }
            
            pairNumber++;
        }

        return string.Join("\n", schedule);
    }

    private static async Task<byte[]> DownloadFile()
    {
        using var httpClient = new HttpClient();
        var response = await httpClient.GetAsync("https://serp-koll.ru/images/ep/k2/rasp2.xlsx");
        return await response.Content.ReadAsByteArrayAsync();
    }

    private static async Task<string> DownloadAndSaveFileAsync(string url, string directory, string fileName = null)
    {
        try
        {
            using var httpClient = new HttpClient();
            var response = await httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            var fileBytes = await response.Content.ReadAsByteArrayAsync();

            var tempPath = Path.Combine(directory, "temp.xlsx");
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            await System.IO.File.WriteAllBytesAsync(tempPath, fileBytes);

            if (string.IsNullOrEmpty(fileName))
            {
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(tempPath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet != null)
                    {
                        var dateCell = worksheet.Cells["A1"].Text?.Trim() ?? "";
                        if (string.IsNullOrEmpty(dateCell))
                            dateCell = worksheet.Cells["B1"].Text?.Trim() ?? "";

                        if (!string.IsNullOrEmpty(dateCell))
                        {
                            fileName = $"Schedule_{DateTime.Now:dd_MM_yyyy_HH_mm}.xlsx";
                        }
                        else
                        {
                            fileName = $"Schedule_{DateTime.Now:dd_MM_yyyy_HH_mm}.xlsx";
                        }
                    }
                }

                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = $"Schedule_{DateTime.Now:dd_MM_yyyy_HH_mm}.xlsx";
                }
            }

            var fullPath = Path.Combine(directory, fileName);

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

    private static string GetPreviousSchedulePath(long chatId)
    {
        var scheduleFiles = Directory.GetFiles(directoryPath, "Schedule_*.xlsx");
        if (scheduleFiles.Length < 2) return null;

        return scheduleFiles.OrderByDescending(f => new FileInfo(f).CreationTime).Skip(1).FirstOrDefault();
    }

    public class ScheduleBlock
    {
        public string DateCell { get; set; }
        public int GroupRow { get; set; }
        public int GroupStartCol { get; set; }
        public int GroupEndCol { get; set; }
        public int TimeStartRow { get; set; }
        public int TimeCol { get; set; }
        public int ScheduleStartRow { get; set; }
        public int ScheduleEndRow { get; set; }
        public int ScheduleStartCol { get; set; }
        public int ScheduleEndCol { get; set; }
        public bool HasClassHour { get; set; }
    }
}
