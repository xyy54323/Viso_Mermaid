using System;
using System.Windows.Forms;

namespace VisioAddIn1
{
    internal static class UserNotificationService
    {
        private const string ErrorTitle = "错误";
        private const string InfoTitle = "提示";
        private const string SuccessTitle = "成功";

        public static void ShowMissingApplication()
        {
            ShowError("无法获取Visio应用程序对象，无法继续操作");
        }

        public static void ShowError(string message)
        {
            MessageBox.Show(message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void ShowError(string prefix, Exception exception)
        {
            if (exception == null)
            {
                ShowError(prefix);
                return;
            }

            ShowError($"{prefix}: {exception.Message}");
        }

        public static void ShowDetailedError(string prefix, Exception exception)
        {
            if (exception == null)
            {
                ShowError(prefix);
                return;
            }

            ShowError($"{prefix}: {exception.Message}\n\n{exception.StackTrace}");
        }

        public static void ShowInfo(string message)
        {
            MessageBox.Show(message, InfoTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ShowSuccess(string message)
        {
            MessageBox.Show(message, SuccessTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
