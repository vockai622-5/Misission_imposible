using System;

namespace macros.Services
{
    public class PasswordGenerator
    {
        private readonly Random _random;

        public PasswordGenerator()
        {
            _random = new Random();
        }

        public string GeneratePin()
        {
            return _random.Next(10000, 100000).ToString();
        }
    }
}