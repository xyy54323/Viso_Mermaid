using System;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    internal static class MermaidFlowchartService
    {
        public static void Execute(Visio.Application visioApp)
        {
            try
            {
                if (visioApp == null)
                {
                    UserNotificationService.ShowMissingApplication();
                    return;
                }

                using (var form = new MermaidForm())
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    var flowchartData = form.ParsedFlowchartData;
                    if (flowchartData == null)
                    {
                        var parser = new MermaidParser();
                        flowchartData = parser.Parse(form.MermaidCode);
                    }

                    var generator = new VisioFlowchartGenerator(visioApp);
                    generator.GenerateFlowchart(flowchartData);

                    UserNotificationService.ShowSuccess("流程图已成功生成！");
                }
            }
            catch (Exception ex)
            {
                UserNotificationService.ShowDetailedError("生成流程图时出错", ex);
            }
        }
    }
}
