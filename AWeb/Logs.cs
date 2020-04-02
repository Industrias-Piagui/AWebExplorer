using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace AWeb
{
    public static class Logs
    {
        private static readonly string LogPath;
        private static readonly string ErrorLogPath;

        static Logs()
        {
            LogPath = $@"{Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)}\Logs\Default.log";
            ErrorLogPath = $@"{Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)}\Logs\Error.log";
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void WriteLog(string message)
        {
            var now = DateTime.Now;
            File.AppendAllText(LogPath, $"[{now}] - {message}\r\n");
        }

        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void WriteErrorLog(string message)
        {
            var now = DateTime.Now;
            File.AppendAllText(ErrorLogPath, $"[{now}] - {message}\r\n");
        }
    }
}
