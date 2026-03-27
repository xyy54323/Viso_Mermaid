using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    public class VisioFlowchartGenerator
    {
        private const string DefaultStencilName = "BASIC_U.VSS";
        private const string FallbackStencilName = "BASIC_M.VSS";

        private readonly Visio.Application _application;
        private readonly VisioFlowchartShapeFactory _shapeFactory;
        private readonly VisioFlowchartLayoutEngine _layoutEngine;
        private readonly VisioFlowchartConnectionRouter _connectionRouter;

        public VisioFlowchartGenerator(Visio.Application application)
        {
            _application = application;
            _shapeFactory = new VisioFlowchartShapeFactory(application);
            _layoutEngine = new VisioFlowchartLayoutEngine();
            _connectionRouter = new VisioFlowchartConnectionRouter(application);
        }

        public void GenerateFlowchart(MermaidParser.FlowchartData flowchartData)
        {
            if (flowchartData == null)
            {
                return;
            }

            try
            {
                var page = PrepareActivePage(flowchartData.Direction);
                EnsureFlowchartStencil();

                var shapeMap = CreateShapes(page, flowchartData);
                _layoutEngine.ApplyManualLayout(flowchartData, shapeMap, page);
                _connectionRouter.CreateConnections(page, flowchartData, shapeMap);
                _layoutEngine.AutoLayoutDiagram(page);
                DeselectAll();
            }
            catch (Exception ex)
            {
                InternalLog.Error("生成流程图失败", ex);
            }
        }

        private Visio.Page PrepareActivePage(string direction)
        {
            EnsureActiveDocument();
            var page = _application.ActivePage;
            _layoutEngine.ConfigurePageForDirection(page, direction);
            return page;
        }

        private Dictionary<string, Visio.Shape> CreateShapes(Page page, MermaidParser.FlowchartData flowchartData)
        {
            var shapeMap = new Dictionary<string, Visio.Shape>(StringComparer.Ordinal);

            foreach (var node in flowchartData.Nodes)
            {
                if (!IsRenderableNode(node))
                {
                    continue;
                }

                var shape = _shapeFactory.CreateShape(page, node);
                if (shape != null)
                {
                    shapeMap[node.Id] = shape;
                }
            }

            return shapeMap;
        }

        private bool IsRenderableNode(MermaidParser.Node node)
        {
            return node != null &&
                   !string.IsNullOrWhiteSpace(node.Id) &&
                   !string.Equals(node.Id, "-", StringComparison.Ordinal);
        }

        private Document EnsureActiveDocument()
        {
            try
            {
                return _application.ActiveDocument;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"当前没有活动文档，改为新建文档: {ex.Message}");
                return _application.Documents.Add(string.Empty);
            }
        }

        private void EnsureFlowchartStencil()
        {
            try
            {
                if (TryOpenStencil(ResolveStencilCandidates()))
                {
                    return;
                }

                TryOpenStencil(new[] { FallbackStencilName });
            }
            catch (Exception ex)
            {
                InternalLog.Error("准备流程图模具失败，将退回基本形状", ex);
            }
        }

        private IEnumerable<string> ResolveStencilCandidates()
        {
            yield return DefaultStencilName;

            string visioPath = null;
            try
            {
                visioPath = _application.Path;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"读取 Visio 安装路径失败，改用默认模具名: {ex.Message}");
                yield break;
            }

            var relativeCandidates = new[]
            {
                @"1033\VISIO\BASIC_U.VSS",
                @"VISIO\BASIC_U.VSS",
                @"BASIC_U.VSS",
                @"..\BASIC_U.VSS",
                @"..\VISIO\BASIC_U.VSS",
                @"..\..\VISIO\BASIC_U.VSS"
            };

            foreach (var relativePath in relativeCandidates)
            {
                string fullPath = System.IO.Path.Combine(visioPath, relativePath);
                if (System.IO.File.Exists(fullPath))
                {
                    yield return fullPath;
                }
            }
        }

        private bool TryOpenStencil(IEnumerable<string> stencilCandidates)
        {
            foreach (var stencilCandidate in stencilCandidates)
            {
                try
                {
                    _application.Documents.OpenEx(stencilCandidate, (short)Visio.VisOpenSaveArgs.visOpenDocked);
                    return true;
                }
                catch (Exception ex)
                {
                    InternalLog.Info($"打开模具失败，尝试下一个: {stencilCandidate} ({ex.Message})");
                }
            }

            return false;
        }

        private void DeselectAll()
        {
            try
            {
                _application.ActiveWindow.DeselectAll();
            }
            catch (Exception ex)
            {
                InternalLog.Info($"取消选择失败: {ex.Message}");
            }
        }
    }
}
