using System;
using System.Collections.Generic;
using System.IO;

namespace SerpThunder
{
    class Format
    {
        private readonly string _filePath;
        private Dictionary<long, string> _chatFormats;

        public Format(string filePath)
        {
            _filePath = filePath;
            _chatFormats = new Dictionary<long, string>();
            LoadFormats();
        }

        private void LoadFormats()
        {
            if (!File.Exists(_filePath))
                return;

            foreach (var line in File.ReadAllLines(_filePath))
            {
                var parts = line.Split('|');
                if (parts.Length == 2 && long.TryParse(parts[0], out var chatId))
                {
                    _chatFormats[chatId] = parts[1];
                }
            }
        }

        private void SaveFormats()
        {
            using var writer = new StreamWriter(_filePath);
            foreach (var chatFormat in _chatFormats)
            {
                writer.WriteLine($"{chatFormat.Key}|{chatFormat.Value}");
            }
        }

        public void SetFormat(long chatId, string format)
        {
            if (string.IsNullOrWhiteSpace(format))
                throw new ArgumentException("Format cannot be null or empty.", nameof(format));

            _chatFormats[chatId] = format;
            SaveFormats();
        }

        public string GetFormat(long chatId)
        {
            return _chatFormats.TryGetValue(chatId, out var format) ? format : "text"; // Default to text
        }

        public void RemoveFormat(long chatId)
        {
            if (_chatFormats.Remove(chatId))
            {
                SaveFormats();
            }
        }

        public Dictionary<long, string> GetAllFormats()
        {
            return new Dictionary<long, string>(_chatFormats);
        }
    }
}
