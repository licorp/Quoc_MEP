using System;
using System.IO;

namespace Quoc_MEP
{
    /// <summary>
    /// Logger cho MEPToolsPanel
    /// </summary>
    public static class MEPToolsPanelLogger
    {
        private static readonly string LogFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "RevitMEP_Panel.log"
        );

        /// <summary>
        /// Log thông tin
        /// </summary>
        public static void Info(string message)
        {
            WriteLog("INFO", message);
        }

        /// <summary>
        /// Log lỗi
        /// </summary>
        public static void Error(string message, Exception ex = null)
        {
            if (ex != null)
            {
                WriteLog("ERROR", $"{message}: {ex.Message}\n{ex.StackTrace}");
            }
            else
            {
                WriteLog("ERROR", message);
            }
        }

        /// <summary>
        /// Log cảnh báo
        /// </summary>
        public static void Warning(string message)
        {
            WriteLog("WARNING", message);
        }

        /// <summary>
        /// Log bắt đầu thao tác
        /// </summary>
        public static void StartOperation(string operationName)
        {
            WriteLog("START", $"Operation: {operationName}");
        }

        /// <summary>
        /// Log kết thúc thao tác
        /// </summary>
        public static void EndOperation(string operationName)
        {
            WriteLog("END", $"Operation: {operationName}");
        }

        /// <summary>
        /// Ghi log vào file
        /// </summary>
        private static void WriteLog(string level, string message)
        {
            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                File.AppendAllText(LogFile, logMessage + Environment.NewLine);
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }
}
