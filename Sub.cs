using System;
using System.Collections.Generic;
using System.IO;

namespace SerpThunder
{
    public class Sub
    {
        private readonly string _filePath;
        private Dictionary<long, (string type, string name)> _subscriptions;
        private Dictionary<long, string> _userStates;

        public Sub(string filePath)
        {
            _filePath = filePath;
            _subscriptions = new Dictionary<long, (string type, string name)>();
            _userStates = new Dictionary<long, string>();
        }

        public void LoadSubscriptions()
        {
            if (!File.Exists(_filePath))
                return;

            foreach (var line in File.ReadAllLines(_filePath))
            {
                var parts = line.Split('|');
                if (parts.Length == 3 && long.TryParse(parts[0], out var chatId))
                {
                    _subscriptions[chatId] = (parts[1], parts[2]);
                }
            }
        }

        public void SaveSubscriptions()
        {
            using var writer = new StreamWriter(_filePath);
            foreach (var subscription in _subscriptions)
            {
                writer.WriteLine($"{subscription.Key}|{subscription.Value.type}|{subscription.Value.name}");
            }
        }

        public void AddSubscription(long chatId, string type, string name)
        {
            _subscriptions[chatId] = (type, name); // Добавляем или обновляем подписку
            SaveSubscriptions(); // Сохраняем в файл
        }

        public bool TryGetSubscription(long chatId, out (string type, string name) subscription)
        {
            return _subscriptions.TryGetValue(chatId, out subscription);
        }

        public void RemoveSubscription(long chatId)
        {
            if (_subscriptions.Remove(chatId))
            {
                SaveSubscriptions();
            }
        }

        public Dictionary<long, (string type, string name)> GetAllSubscriptions()
        {
            return new Dictionary<long, (string type, string name)>(_subscriptions);
        }

        public void SetState(long chatId, string state)
        {
            _userStates[chatId] = state;
        }

        public bool TryGetState(long chatId, out string state)
        {
            return _userStates.TryGetValue(chatId, out state);
        }

        public void ClearState(long chatId)
        {
            _userStates.Remove(chatId);
        }
    }
}
