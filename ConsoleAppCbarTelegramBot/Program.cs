namespace ConsoleAppCbarTelegramBot
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            var telegramHelper = new TelegramBotHelper("6883678634:AAHgFZhWy175MKvA44wy2n9OwspkGCAqE6s");
            telegramHelper.GetUpdates();
        }
    }
}