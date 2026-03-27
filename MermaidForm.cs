using System;
using System.Windows.Forms;

namespace VisioAddIn1
{
    public partial class MermaidForm : Form
    {
        public string MermaidCode { get; private set; }
        public MermaidParser.FlowchartData ParsedFlowchartData { get; private set; }

        public MermaidForm()
        {
            InitializeComponent();
            
            // 确保文本框可以接收输入
            txtMermaidCode.ReadOnly = false;
            txtMermaidCode.Enabled = true;
            
            // 设置默认示例代码
            txtMermaidCode.Text = "graph TD\r\nA[开始] --> B{条件}\r\nB -->|是| C[处理1]\r\nB -->|否| D[处理2]\r\nC --> E[结束]\r\nD --> E";
            
            // 设置焦点到文本框
            this.Load += (s, e) => { txtMermaidCode.Focus(); };
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                MermaidCode = txtMermaidCode.Text;
                
                if (string.IsNullOrWhiteSpace(MermaidCode))
                {
                    UserNotificationService.ShowInfo("请输入Mermaid流程图代码");
                    return;
                }
                
                // 解析Mermaid代码（不再显示调试信息）
                var parser = new MermaidParser();
                ParsedFlowchartData = parser.Parse(MermaidCode);

                // 直接关闭窗口，返回OK结果
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                UserNotificationService.ShowError("处理输入时出错", ex);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
        
        // 添加键盘快捷键支持
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Enter))
            {
                btnGenerate_Click(this, EventArgs.Empty);
                return true;
            }
            else if (keyData == Keys.Escape)
            {
                btnCancel_Click(this, EventArgs.Empty);
                return true;
            }
            
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
