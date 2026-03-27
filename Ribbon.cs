using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private Visio.Application _application;

        public Ribbon(Visio.Application application)
        {
            _application = application;

            if (_application == null)
            {
                BeginDeferredApplicationInitialization();
            }
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' loadImage='LoadImage' onLoad='OnLoad'>
              <ribbon>
                <tabs>
                  <tab id='CustomTab' label='Mermaid流程图'>
                    <group id='MermaidGroup' label='Mermaid工具'>
                      <button id='MermaidButton' 
                              label='从Mermaid代码生成流程图' 
                              size='large'
                              onAction='OnMermaidButtonClick'/>
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        #endregion

        #region 功能区回调

        public void OnLoad(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            TryResolveApplication(out _application);
        }

        public void OnMermaidButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (!TryResolveApplication(out var visioApp))
                {
                    UserNotificationService.ShowMissingApplication();
                    return;
                }

                MermaidFlowchartService.Execute(visioApp);
            }
            catch (Exception ex)
            {
                UserNotificationService.ShowError("生成流程图时出错", ex);
            }
        }

        public System.Drawing.Bitmap LoadImage(string imageName)
        {
            // 这里可以加载自定义图标，暂时返回null
            return null;
        }

        private void BeginDeferredApplicationInitialization()
        {
            try
            {
                System.Threading.Timer timer = null;
                timer = new System.Threading.Timer(state =>
                {
                    try
                    {
                        if (TryResolveApplication(out var resolvedApplication))
                        {
                            _application = resolvedApplication;
                            ribbon?.Invalidate();
                            timer.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        InternalLog.Error("延迟初始化出错", ex);
                    }
                }, null, 500, 500);
            }
            catch (Exception ex)
            {
                InternalLog.Error("尝试获取Visio应用程序对象时出错", ex);
            }
        }

        private bool TryResolveApplication(out Visio.Application application)
        {
            application = _application;
            if (application != null)
            {
                return true;
            }

            try
            {
                if (Globals.ThisAddIn != null)
                {
                    application = Globals.ThisAddIn.Application;
                    if (application != null)
                    {
                        _application = application;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                InternalLog.Error("获取全局Visio应用程序对象时出错", ex);
            }

            application = null;
            return false;
        }

        #endregion
    }
}
