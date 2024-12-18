using OfficeOpenXml;
using SkiaSharp;

namespace SerpThunder
{
    class Converter
    {
        public static SKPaint textPaint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 14,
            IsAntialias = true,
            TextAlign = SKTextAlign.Left,
            Typeface = SKTypeface.FromFamilyName("Ariel")
        };

        public static SKPaint boldTextPaint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 20,
            IsAntialias = true,
            TextAlign = SKTextAlign.Left,
            Typeface = SKTypeface.FromFamilyName("Ariel", SKFontStyle.Bold)
        };
        public static SKPaint LeftTextPaint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 40,
            IsAntialias = true,
            TextAlign = SKTextAlign.Left,
            Typeface = SKTypeface.FromFamilyName("Ariel", SKFontStyle.Bold)
        };

        public static SKPaint TopTextPaint = new SKPaint
        {
            Color = SKColors.Black,
            TextSize = 26,
            IsAntialias = true,
            TextAlign = SKTextAlign.Left,
            Typeface = SKTypeface.FromFamilyName("Ariel", SKFontStyle.Bold)
        };
        public static string[,] ParseScheduleToMatrix(string scheduleText, string groupName)
        {
            if (string.IsNullOrWhiteSpace(scheduleText))
                throw new ArgumentException("Текст расписания не может быть пустым.", nameof(scheduleText));

            var endOfFirstLine = scheduleText.IndexOf('\n');
            if (endOfFirstLine >= 0)
            {
                scheduleText = scheduleText.Substring(endOfFirstLine + 1);
            }

            var lines = scheduleText.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var scheduleMatrix = new string[6, 2];
            scheduleMatrix[0, 1] = groupName;

            var currentPairData = string.Empty;
            var currentPairIndex = 0;

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                if (trimmedLine.Length >= 2 && char.IsDigit(trimmedLine[0]) && trimmedLine[1] == ')')
                {
                    if (currentPairIndex > 0 && currentPairIndex <= 5)
                    {
                        scheduleMatrix[currentPairIndex, 1] = ProcessPairData(currentPairData);
                    }

                    currentPairIndex = int.Parse(trimmedLine[0].ToString());
                    currentPairData = trimmedLine.Substring(2).Trim();
                }
                else
                {
                    currentPairData += "\n" + trimmedLine;
                }
            }

            if (currentPairIndex > 0 && currentPairIndex <= 5)
            {
                scheduleMatrix[currentPairIndex, 1] = ProcessPairData(currentPairData);
            }

            for (int i = 1; i <= 5; i++)
            {
                scheduleMatrix[i, 0] = i.ToString();
            }

            return scheduleMatrix;
        }

        private static string ProcessPairData(string pairData)
        {
            if (string.IsNullOrWhiteSpace(pairData) || pairData == "Нет пары")
                return null;

            return pairData
                .Replace("Преподаватель: ", string.Empty)
                .Replace("Кабинет: ", string.Empty)
                .Replace("Разделённая пара:", string.Empty)
                .Replace("Подгруппа 1:", string.Empty)
                .Replace("Подгруппа 2:", "|")
                .Trim();
        }

        public static void ConvertScheduleToImage(string groupName, string scheduleText, string outputPath, string filePath)
        {
            // Читаем дату и день недели из Excel
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets.First();

            string dateRaw = worksheet.Cells[1, 2].Text.Trim();
            string prefix = "Расписание на ";
            string dateOnly = dateRaw.StartsWith(prefix) ? dateRaw.Substring(prefix.Length) : dateRaw;

            string dayOfWeek = worksheet.Cells[1, 16].Text.Trim();

            var scheduleMatrix = ParseScheduleToMatrix(scheduleText, groupName);

            int leftCellWidth = 35;
            int rightCellWidth = 140;
            int cellHeight = 300;
            int headerHeight = 25;
            int dateHeight = 30; // Высота строки с датой
            int dayOfWeekHeight = 25; // Высота строки с днём недели
            int imageWidth = leftCellWidth + rightCellWidth;
            int imageHeight = dateHeight + dayOfWeekHeight + headerHeight + cellHeight * 5;

            using var surface = SKSurface.Create(new SKImageInfo(imageWidth, imageHeight));
            var canvas = surface.Canvas;
            canvas.Clear(SKColors.White);

            var pen = new SKPaint
            {
                Style = SKPaintStyle.Stroke,
                Color = SKColors.Black,
                StrokeWidth = 2
            };

            // Рисуем дату
            var dateRect = new SKRect(-17, 0, imageWidth, dateHeight);
            DrawHeaderCenteredText(canvas, dateOnly, dateRect, boldTextPaint);

            // Рисуем день недели
            var dayOfWeekRect = new SKRect(-17, dateHeight, imageWidth, dateHeight + dayOfWeekHeight);
            DrawHeaderCenteredText(canvas, dayOfWeek, dayOfWeekRect, textPaint);

            // Рисуем заголовок (группу)
            var headerRect = new SKRect(0, dateHeight + dayOfWeekHeight, imageWidth, dateHeight + dayOfWeekHeight + headerHeight);
            canvas.DrawRect(headerRect, pen);
            DrawHeaderCenteredText(canvas, groupName, headerRect, boldTextPaint);

            canvas.DrawLine(leftCellWidth, dateHeight + dayOfWeekHeight, leftCellWidth, dateHeight + dayOfWeekHeight + headerHeight, pen);

            for (int i = 0; i < 5; i++)
            {
                float top = dateHeight + dayOfWeekHeight + headerHeight + i * cellHeight;

                // Left cell
                var leftRect = new SKRect(0, top, leftCellWidth, top + cellHeight);
                canvas.DrawRect(leftRect, pen);

                if (!string.IsNullOrEmpty(scheduleMatrix[i + 1, 0]))
                {
                    DrawCenteredNumber(canvas, scheduleMatrix[i + 1, 0], leftRect, LeftTextPaint);
                }

                // Right cell
                var rightRect = new SKRect(leftCellWidth, top, imageWidth, top + cellHeight);
                canvas.DrawRect(rightRect, pen);

                if (!string.IsNullOrEmpty(scheduleMatrix[i + 1, 1]))
                {
                    var dataParts = scheduleMatrix[i + 1, 1].Split('|');

                    if (dataParts.Length == 2)
                    {
                        var topTextLines = FormatText(dataParts[0], textPaint, rightCellWidth - 10);
                        var bottomTextLines = FormatText(dataParts[1], textPaint, rightCellWidth - 10);

                        float topHalfHeight = cellHeight / 2;
                        float topStartY = top + (topHalfHeight - topTextLines.Count * (textPaint.TextSize + 5)) / 2;

                        foreach (var line in topTextLines)
                        {
                            var textRect = new SKRect(leftCellWidth, topStartY, imageWidth, topStartY + textPaint.TextSize);
                            DrawCenteredText(canvas, line, textRect, textPaint);
                            topStartY += textPaint.TextSize + 5;
                        }

                        float bottomStartY = top + topHalfHeight + (topHalfHeight - bottomTextLines.Count * (textPaint.TextSize + 5)) / 2;

                        foreach (var line in bottomTextLines)
                        {
                            var textRect = new SKRect(leftCellWidth, bottomStartY, imageWidth, bottomStartY + textPaint.TextSize);
                            DrawCenteredText(canvas, line, textRect, textPaint);
                            bottomStartY += textPaint.TextSize + 5;
                        }

                        canvas.DrawLine(leftCellWidth, top + topHalfHeight, imageWidth, top + topHalfHeight, pen);
                    }
                    else
                    {
                        var dataLines = FormatText(scheduleMatrix[i + 1, 1], textPaint, rightCellWidth - 10);
                        float totalTextHeight = dataLines.Count * (textPaint.TextSize + 5);
                        float startY = rightRect.Top + (cellHeight - totalTextHeight) / 2;

                        foreach (var line in dataLines)
                        {
                            var textRect = new SKRect(leftCellWidth, startY, imageWidth, startY + textPaint.TextSize);
                            DrawCenteredText(canvas, line, textRect, textPaint);
                            startY += textPaint.TextSize + 5;
                        }
                    }
                }
            }

            using var image = surface.Snapshot();
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            using var stream = File.OpenWrite(outputPath);
            data.SaveTo(stream);
        }



        private static List<string> FormatText(string text, SKPaint paint, float maxWidth)
        {
            var lines = text.Split('\n');
            var formattedLines = new List<string>();

            foreach (var line in lines)
            {
                var words = line.Split(' ');
                var currentLine = "";

                foreach (var word in words)
                {
                    var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
                    if (paint.MeasureText(testLine) > maxWidth)
                    {
                        formattedLines.Add(currentLine);
                        currentLine = word;
                    }
                    else
                    {
                        currentLine = testLine;
                    }
                }

                if (!string.IsNullOrEmpty(currentLine))
                {
                    formattedLines.Add(currentLine);
                }
            }

            return formattedLines;
        }

        private static void DrawCenteredText(SKCanvas canvas, string text, SKRect rect, SKPaint paint)
        {
            // Измеряем границы текста
            var bounds = new SKRect();
            paint.MeasureText(text, ref bounds);

            // Расчёт координат для центрирования
            float x = rect.MidX - (bounds.Width / 2);
            float y = rect.MidY + (paint.TextSize / 3); // Корректировка для вертикального выравнивания

            // Рисуем текст
            canvas.DrawText(text, x, y, paint);
        }
        private static void DrawCenteredNumber(SKCanvas canvas, string text, SKRect rect, SKPaint paint)
        {
            // Измеряем границы текста
            var bounds = new SKRect();
            paint.MeasureText(text, ref bounds);

            // Расчёт координат для центрирования
            float x = rect.MidX - (bounds.Width / 2)-2;
            float y = rect.MidY + (paint.TextSize / 3); // Корректировка для вертикального выравнивания

            // Рисуем текст
            canvas.DrawText(text, x, y, paint);
        }
        private static void DrawHeaderCenteredText(SKCanvas canvas, string text, SKRect rect, SKPaint paint)
        {            
            var paintToUse = text.Trim().All(char.IsDigit) ? TopTextPaint : boldTextPaint;
            // Измеряем границы текста
            var bounds = new SKRect();
            paintToUse.MeasureText(text, ref bounds);

            // Расчёт координат для центрирования
            float x = text.Trim().All(char.IsDigit) ? rect.MidX - (bounds.Width / 2)+16: rect.MidX - (bounds.Width / 2) + 17;
            float y = rect.MidY + (paintToUse.TextSize / 3); // Корректировка для вертикального выравнивания

            // Рисуем текст
            canvas.DrawText(text, x, y, paintToUse);
        }
    }
}
