using System;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    internal sealed class VisioFlowchartShapeFactory
    {
        private const string RectangleShape = "rectangle";
        private const string RoundedRectangleShape = "rounded rectangle";
        private const string DiamondShape = "diamond";
        private const string DatabaseShape = "database";
        private const string CircleShape = "circle";

        private const double DefaultShapeSize = 1.0;
        private const double StencilDropX = 2.0;
        private const double StencilDropY = 2.0;
        private const double RoundedCornerAmount = 0.1;

        private const double DefaultMinWidth = 1.0;
        private const double DefaultMaxWidth = 5.0;
        private const double DefaultWidthScale = 0.12;
        private const double DefaultWidthPadding = 0.4;

        private const double BoxMinWidth = 0.85;
        private const double BoxMaxWidth = 4.2;
        private const double BoxWidthScale = 0.1;
        private const double BoxWidthPadding = 0.25;

        private const double MinHeight = 0.6;
        private const double MaxHeight = 2.6;
        private const double HeightScale = 0.24;
        private const double HeightPadding = 0.28;

        private readonly Visio.Application _application;

        public VisioFlowchartShapeFactory(Visio.Application application)
        {
            _application = application;
        }

        public Visio.Shape CreateShape(Visio.Page page, MermaidParser.Node node)
        {
            string nodeText = GetNodeText(node);
            string shapeType = node != null ? node.Shape : null;

            try
            {
                return BuildShape(page, shapeType, nodeText, useStencilShape: true);
            }
            catch (Exception ex)
            {
                InternalLog.Info($"优先形状创建失败，尝试回退形状: {shapeType} ({ex.Message})");
                try
                {
                    return BuildShape(page, shapeType, nodeText, useStencilShape: false);
                }
                catch (Exception fallbackEx)
                {
                    InternalLog.Error($"回退形状创建失败: {shapeType}", fallbackEx);
                    return null;
                }
            }
        }

        private Visio.Shape BuildShape(Visio.Page page, string shapeType, string nodeText, bool useStencilShape)
        {
            var shape = useStencilShape
                ? CreatePreferredShape(page, shapeType)
                : CreateFallbackShape(page, shapeType);

            shape.Text = nodeText;
            SetShapeStyle(shape);
            AdjustShapeSize(shape, shapeType);
            SetTextAlignment(shape);
            return shape;
        }

        private string GetNodeText(MermaidParser.Node node)
        {
            return string.IsNullOrWhiteSpace(node.Text) ? node.Id : node.Text;
        }

        private Visio.Shape CreatePreferredShape(Visio.Page page, string shapeType)
        {
            switch (shapeType)
            {
                case DiamondShape:
                    return TryDropStencilShape(page, "Decision", "Diamond", "Condition") ?? CreateDiamond(page);
                case DatabaseShape:
                    return TryDropStencilShape(page, "Database", "Data", "Cylinder") ?? CreateRectangle(page);
                case CircleShape:
                    return page.DrawOval(0, 0, DefaultShapeSize, DefaultShapeSize);
                case RoundedRectangleShape:
                    return CreateRoundedRectangle(page);
                default:
                    return CreateRectangle(page);
            }
        }

        private Visio.Shape CreateFallbackShape(Visio.Page page, string shapeType)
        {
            switch (shapeType)
            {
                case DiamondShape:
                    return CreateDiamond(page);
                case CircleShape:
                    return page.DrawOval(0, 0, DefaultShapeSize, DefaultShapeSize);
                case RoundedRectangleShape:
                    return CreateRoundedRectangle(page);
                default:
                    return CreateRectangle(page);
            }
        }

        private Visio.Shape CreateRectangle(Visio.Page page)
        {
            return page.DrawRectangle(0, 0, DefaultShapeSize, DefaultShapeSize);
        }

        private Visio.Shape CreateRoundedRectangle(Visio.Page page)
        {
            var rectangle = CreateRectangle(page);
            rectangle.CellsU["Rounding"].Formula = RoundedCornerAmount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return rectangle;
        }

        private Visio.Shape TryDropStencilShape(Visio.Page page, params string[] keywords)
        {
            foreach (Visio.Document document in _application.Documents)
            {
                if (!IsBasicStencil(document))
                {
                    continue;
                }

                foreach (Visio.Master master in document.Masters)
                {
                    if (keywords.Any(keyword => master.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        return page.Drop(master, StencilDropX, StencilDropY);
                    }
                }
            }

            return null;
        }

        private bool IsBasicStencil(Visio.Document document)
        {
            return document.Name.IndexOf("BASIC", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   document.Name.EndsWith(".VSS", StringComparison.OrdinalIgnoreCase);
        }

        private Visio.Shape CreateDiamond(Visio.Page page)
        {
            double[] xyArray =
            {
                DefaultShapeSize / 2, 0,
                DefaultShapeSize, DefaultShapeSize / 2,
                DefaultShapeSize / 2, DefaultShapeSize,
                0, DefaultShapeSize / 2,
                DefaultShapeSize / 2, 0
            };

            return page.DrawSpline(xyArray, 0, (short)Visio.VisDrawSplineFlags.visSplineAbrupt);
        }

        private void SetShapeStyle(Visio.Shape shape)
        {
            try
            {
                shape.CellsU["FillPattern"].Formula = "0";
                shape.CellsU["LineColor"].FormulaForceU = "THEMEGUARD(RGB(0,0,0))";
                shape.CellsU["LinePattern"].Formula = "1";
                shape.CellsU["LineWeight"].Formula = "1.5 pt";
                shape.CellsU["Char.Color"].FormulaForceU = "THEMEGUARD(RGB(0,0,0))";
                shape.CellsU["Char.Style"].Formula = "0";
                shape.CellsU["Char.Size"].Formula = "12 pt";
            }
            catch (Exception ex)
            {
                InternalLog.Info($"设置节点样式失败: {ex.Message}");
            }
        }

        private void AdjustShapeSize(Visio.Shape shape, string shapeType)
        {
            try
            {
                var sizeProfile = MeasureText(shape.Text);
                if (sizeProfile.lineCount == 0)
                {
                    return;
                }

                double centerX = shape.CellsU["PinX"].ResultIU;
                double centerY = shape.CellsU["PinY"].ResultIU;

                shape.CellsU["Width"].ResultIU = CalculateWidth(sizeProfile.maxLineLength, shapeType);
                shape.CellsU["Height"].ResultIU = CalculateHeight(sizeProfile.lineCount);
                shape.CellsU["PinX"].ResultIU = centerX;
                shape.CellsU["PinY"].ResultIU = centerY;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"调整节点大小失败: {ex.Message}");
            }
        }

        private (int maxLineLength, int lineCount) MeasureText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return (0, 0);
            }

            int maxLineLength = 0;
            string[] lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                maxLineLength = Math.Max(maxLineLength, GetVisualTextLength(line));
            }

            return (maxLineLength, lines.Length);
        }

        private double CalculateWidth(int maxLineLength, string shapeType)
        {
            if (IsBoxLikeShape(shapeType))
            {
                return Clamp(maxLineLength * BoxWidthScale + BoxWidthPadding, BoxMinWidth, BoxMaxWidth);
            }

            return Clamp(maxLineLength * DefaultWidthScale + DefaultWidthPadding, DefaultMinWidth, DefaultMaxWidth);
        }

        private bool IsBoxLikeShape(string shapeType)
        {
            return string.Equals(shapeType, RectangleShape, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(shapeType, RoundedRectangleShape, StringComparison.OrdinalIgnoreCase);
        }

        private double CalculateHeight(int lineCount)
        {
            return Clamp(lineCount * HeightScale + HeightPadding, MinHeight, MaxHeight);
        }

        private double Clamp(double value, double min, double max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private int GetVisualTextLength(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int length = 0;
            foreach (char ch in text)
            {
                length += ch <= 127 ? 1 : 2;
            }

            return length;
        }

        private void SetTextAlignment(Visio.Shape shape)
        {
            try
            {
                shape.CellsU["TextXForm.TxtPinX"].Formula = "Width*0.5";
                shape.CellsU["TextXForm.TxtPinY"].Formula = "Height*0.5";
                shape.CellsU["TextXForm.TxtWidth"].Formula = "Width*1.0";
                shape.CellsU["TextXForm.TxtHeight"].Formula = "Height*1.0";
                shape.CellsU["TextXForm.TxtLocPinX"].Formula = "Width*0.5";
                shape.CellsU["TextXForm.TxtLocPinY"].Formula = "Height*0.5";
                shape.CellsU["Para.HorzAlign"].Formula = "1";
                shape.CellsU["Para.VertAlign"].Formula = "1";
                shape.CellsU["TextXForm.TxtMargin"].Formula = "0.05 in";
            }
            catch (Exception ex)
            {
                InternalLog.Info($"设置文本对齐失败: {ex.Message}");
            }
        }
    }
}
