using System;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    public partial class ThisAddIn
    {
        private Ribbon ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                EnsureRibbonCreated();
            }
            catch (Exception ex)
            {
                UserNotificationService.ShowError("插件启动时出错", ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ribbon = null;
        }

        public static void ShowMermaidFormMacro()
        {
            MermaidFlowchartService.Execute(GetGlobalApplication());
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                var currentRibbon = EnsureRibbonCreated();
                if (currentRibbon != null)
                {
                    return currentRibbon;
                }

                ribbon = ribbon ?? new Ribbon(null);
                return ribbon;
            }
            catch (Exception ex)
            {
                UserNotificationService.ShowError("创建Ribbon对象时出错", ex);
                return null;
            }
        }

        private Ribbon EnsureRibbonCreated()
        {
            if (ribbon == null && Application != null)
            {
                ribbon = new Ribbon(Application);
            }

            return ribbon;
        }

        private static Visio.Application GetGlobalApplication()
        {
            try
            {
                return Globals.ThisAddIn != null ? Globals.ThisAddIn.Application : null;
            }
            catch
            {
                return null;
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
