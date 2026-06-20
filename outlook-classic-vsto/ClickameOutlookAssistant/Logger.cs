using System;
using System.IO;

namespace ClickameOutlookAssistant
{
    /// <summary>
    /// Logger local molt bàsic per a depuració. Escriu a
    /// %APPDATA%\ClickameOutlookAssistant\log.txt
    /// No envia res a cap servidor.
    /// </summary>
    public static class Logger
    {
        private static readonly object _lock = new object();

        public static string LogFolder =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ClickameOutlookAssistant");

        public static string LogFilePath => Path.Combine(LogFolder, "log.txt");

        public static void Info(string message) => Write("INFO", message);

        public static void Warn(string message) => Write("WARN", message);

        public static void Error(string message, Exception ex = null)
        {
            var full = ex == null ? message : $"{message} :: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}";
            Write("ERROR", full);
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (_lock)
                {
                    Directory.CreateDirectory(LogFolder);
                    var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}{Environment.NewLine}";
                    File.AppendAllText(LogFilePath, line);
                }
            }
            catch
            {
                // El logging mai no ha de tombar l'add-in.
            }
        }
    }
}
