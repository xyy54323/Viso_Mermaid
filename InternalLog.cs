using System;
using System.Diagnostics;

namespace VisioAddIn1
{
    internal static class InternalLog
    {
        public static void Info(string message)
        {
            if (!string.IsNullOrWhiteSpace(message))
            {
                Debug.WriteLine($"[VisioAddIn1] {message}");
            }
        }

        public static void Error(string message, Exception exception = null)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                message = "Unexpected error";
            }

            if (exception == null)
            {
                Debug.WriteLine($"[VisioAddIn1] ERROR: {message}");
                return;
            }

            Debug.WriteLine($"[VisioAddIn1] ERROR: {message}: {exception.Message}");
            Debug.WriteLine(exception.StackTrace);
        }
    }
}
