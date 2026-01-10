using System;
using System.IO;
using System.Diagnostics;

namespace Quoc_MEP.Export.Utils
{
    public static class NotificationHelper
    {
        private static readonly string LogFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExportPlusAddin", "Logs");
        
        private static readonly string LogFile = Path.Combine(LogFolder, $"ExportPlus_{DateTime.Now:yyyy-MM-dd}.log");
        
        static NotificationHelper()
        {
            try
            {
                Directory.CreateDirectory(LogFolder);
            }
            catch
            {
                // Ignore directory creation errors
            }
        }
        
        public static void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }
        
        public static void LogWarning(string message)
        {
            WriteLog("WARNING", message);
        }
        
        public static void LogError(string message)
        {
            WriteLog("ERROR", message);
        }
        
        public static void LogError(string message, Exception ex)
        {
            string fullMessage = $"{message}\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
            WriteLog("ERROR", fullMessage);
        }
        
        private static void WriteLog(string level, string message)
        {
            try
            {
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                
                File.AppendAllText(LogFile, logEntry + Environment.NewLine);
                
                // Also write to Debug output for development
                Debug.WriteLine(logEntry);
            }
            catch
            {
                // Ignore logging errors to prevent infinite loops
            }
        }
        
        public static void ShowNotification(string title, string message, NotificationType type = NotificationType.Information)
        {
            try
            {
                // Log the notification
                switch (type)
                {
                    case NotificationType.Information:
                        LogInfo($"{title}: {message}");
                        break;
                    case NotificationType.Warning:
                        LogWarning($"{title}: {message}");
                        break;
                    case NotificationType.Error:
                        LogError($"{title}: {message}");
                        break;
                }
                
                // Show Windows notification if possible
                ShowWindowsNotification(title, message, type);
            }
            catch (Exception ex)
            {
                LogError("Failed to show notification", ex);
            }
        }
        
        private static void ShowWindowsNotification(string title, string message, NotificationType type)
        {
            try
            {
                // Fallback to console output
                Console.WriteLine($"[{type}] {title}: {message}");
                
                // Could integrate with Windows 10 Toast Notifications here
                // or use System.Windows.MessageBox for simple notifications
            }
            catch
            {
                // Ignore notification display errors
            }
        }
        
        public static void ClearOldLogs(int daysToKeep = 30)
        {
            try
            {
                if (Directory.Exists(LogFolder))
                {
                    var files = Directory.GetFiles(LogFolder, "ExportPlus_*.log");
                    var cutoffDate = DateTime.Now.AddDays(-daysToKeep);
                    
                    foreach (string file in files)
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.CreationTime < cutoffDate)
                        {
                            try
                            {
                                File.Delete(file);
                                LogInfo($"Deleted old log file: {Path.GetFileName(file)}");
                            }
                            catch (Exception ex)
                            {
                                LogWarning($"Could not delete old log file {Path.GetFileName(file)}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error cleaning old log files", ex);
            }
        }
        
        public static void OpenLogFolder()
        {
            try
            {
                if (Directory.Exists(LogFolder))
                {
                    Process.Start("explorer.exe", LogFolder);
                }
            }
            catch (Exception ex)
            {
                LogError("Could not open log folder", ex);
            }
        }
        
        public static void SendEmailNotification(string toEmail, string subject, string body, bool isHtml = false)
        {
            try
            {
                // This would require email configuration
                // For now, just log the email that would be sent
                LogInfo($"Email notification would be sent to: {toEmail}");
                LogInfo($"Subject: {subject}");
                LogInfo($"Body: {body}");
                
                // TODO: Implement actual email sending using SMTP
                // Could use System.Net.Mail.MailMessage and SmtpClient
            }
            catch (Exception ex)
            {
                LogError("Failed to send email notification", ex);
            }
        }
        
        public static string GetLogFilePath()
        {
            return LogFile;
        }
        
        public static void ExportLogToFile(string destinationPath)
        {
            try
            {
                if (File.Exists(LogFile))
                {
                    File.Copy(LogFile, destinationPath, true);
                    LogInfo($"Log file exported to: {destinationPath}");
                }
            }
            catch (Exception ex)
            {
                LogError("Failed to export log file", ex);
            }
        }
    }
    
    public enum NotificationType
    {
        Information,
        Warning,
        Error,
        Success
    }
}