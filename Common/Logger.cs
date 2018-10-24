using System;
using System.IO;

namespace Netbattle.Common {
    public class Logger : TaskItem {
        private static string _filename;
        private static LogType _minimumLevel;
        private bool _setup;
        private static readonly object LogLock = new object();

        public override void Setup() {
            if (_setup)
                return;

            LastRun = new DateTime();
            Interval = new TimeSpan(0, 0, 2);

            DateTime nowTime = DateTime.UtcNow;
            _filename = "log." + nowTime.Year + nowTime.Month + nowTime.Day + nowTime.Hour + nowTime.Minute + ".txt";
            File.AppendAllText(_filename, "# Log Start at " + nowTime.ToLongDateString() + " - " + nowTime.ToLongTimeString() + Environment.NewLine);
            _setup = true;
        }

        public override void Main() {
            _minimumLevel = LogType.Verbose;
                //(LogType)Enum.Parse(typeof(LogType), Configuration.Settings.General.LogLevel, true);
        }

        public override void Teardown() {
        }

        public static void Log(LogType type, string message) {
            var item = new LogItem { Type = type, Time = DateTime.UtcNow, Message = message };

            lock (LogLock) {
               // File.AppendAllText(_filename, $"{item.Time.ToLongTimeString()} > [{item.Type}] {item.Message}" + Environment.NewLine);
                //ConsoleOutput(item);
            }
        }

        public static void Log(Exception ex) {
            Log(LogType.Error, $"Error occured: {ex.Message}");
            Log(LogType.Debug, ex.StackTrace);

            if (ex.InnerException == null)
                return;

            Log(LogType.Debug, "INNER EXCEPTION:");
            Log(LogType.Debug, ex.InnerException.Message);
            Log(LogType.Debug, ex.InnerException.StackTrace);
        }

        private static void ConsoleOutput(LogItem item) {
            if ((int)item.Type < (int)_minimumLevel)
                return;

            string line = $"{item.Time.ToLongTimeString()} > ";

            switch (item.Type) {
                case LogType.Verbose:
                    line += "&8[Verbose]";
                    break;
                case LogType.Debug:
                    line += "&8[Debug]";
                    break;
                case LogType.Warning:
                    line += "&e[Warning]&f";
                    break;
                case LogType.Error:
                    line += "&4[Error]&f";
                    break;
                case LogType.Chat:
                    line += "&5[Chat]&f";
                    break;
                case LogType.Command:
                    line += "&a[Command]&8";
                    break;
                case LogType.Info:
                    line += "[Info]&f";
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            line += $" {item.Message}";
         //   ColorConvertingConsole.WriteLine(line);
        }
    }
}
