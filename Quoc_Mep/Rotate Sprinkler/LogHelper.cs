using System;
using System.Diagnostics;
using System.IO;

namespace Quoc_MEP
{
    /// <summary>
    /// Helper để log ra file và DebugView
    /// </summary>
    public static class LogHelper
    {
        private static string _logFilePath;
        private static object _lockObj = new object();

        static LogHelper()
        {
            // Log vào folder Temp
            string tempPath = Path.GetTempPath();
            string logFolder = Path.Combine(tempPath, "QuocMEP_Logs");
            
            try
            {
                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _logFilePath = Path.Combine(logFolder, $"AlignSprinkler_{timestamp}.log");
            }
            catch
            {
                // Nếu không tạo được folder, dùng temp trực tiếp
                _logFilePath = Path.Combine(tempPath, $"QuocMEP_AlignSprinkler.log");
            }
        }

        /// <summary>
        /// Log message ra file và DebugView
        /// </summary>
        public static void Log(string message)
        {
            try
            {
                string logMessage = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";

                // Ghi ra DebugView (OutputDebugString)
                Trace.WriteLine(logMessage);

                // Ghi ra file
                lock (_lockObj)
                {
                    File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
                }
            }
            catch
            {
                // Ignore logging errors
            }
        }

        /// <summary>
        /// Lấy đường dẫn file log hiện tại
        /// </summary>
        public static string GetLogPath()
        {
            return _logFilePath;
        }
    }
}
