using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    internal sealed class VisioFlowchartLayoutEngine
    {
        private const double DefaultPageWidth = 11.0;
        private const double DefaultPageHeight = 8.5;
        private const double DefaultPageCenterX = 5.5;
        private const double LayoutTopPadding = 0.3;
        private const double HorizontalGap = 1.0;
        private const double VerticalGap = 0.55;

        private sealed class GraphLayoutData
        {
            public Dictionary<string, int> NodeOrder { get; } = new Dictionary<string, int>(StringComparer.Ordinal);
            public HashSet<string> ValidNodeIds { get; } = new HashSet<string>(StringComparer.Ordinal);
            public Dictionary<string, List<string>> IncomingEdges { get; } = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            public Dictionary<string, List<string>> OutgoingEdges { get; } = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        }

        public void ConfigurePageForDirection(Visio.Page page, string direction)
        {
            try
            {
                page.PageSheet.CellsU["FlowchartStyle"].Formula = "2";
                page.PageSheet.CellsU["PageWidth"].Formula = $"{DefaultPageWidth} in";
                page.PageSheet.CellsU["PageHeight"].Formula = $"{DefaultPageHeight} in";
                page.PageSheet.CellsU["PageScale"].Formula = "1";
                page.PageSheet.CellsU["DrawingScale"].Formula = "1";
            }
            catch (Exception ex)
            {
                InternalLog.Info($"配置页面失败: {ex.Message}");
            }
        }

        public void AutoLayoutDiagram(Visio.Page page)
        {
            try
            {
                page.PageSheet.CellsU["RouteStyle"].Formula = "5";
                page.PageSheet.CellsU["PlaceStyle"].Formula = "1";
                page.PageSheet.CellsU["PlaceDepth"].Formula = "1";
                page.PageSheet.CellsU["AvenueSizeX"].Formula = "1.5 in";
                page.PageSheet.CellsU["AvenueSizeY"].Formula = "0.3 in";
                page.PageSheet.CellsU["LineToNodeX"].Formula = "0.3 in";
                page.PageSheet.CellsU["LineToNodeY"].Formula = "0.3 in";
                page.PageSheet.CellsU["LineToLineX"].Formula = "0.25 in";
                page.PageSheet.CellsU["LineToLineY"].Formula = "0.25 in";
                page.ResizeToFitContents();
            }
            catch (Exception ex)
            {
                InternalLog.Info($"自动布局失败: {ex.Message}");
            }
        }

        public void ApplyManualLayout(
            MermaidParser.FlowchartData flowchartData,
            Dictionary<string, Visio.Shape> shapeMap,
            Visio.Page page)
        {
            if (shapeMap.Count == 0)
            {
                return;
            }

            var layoutData = BuildGraphLayoutData(flowchartData, shapeMap);
            var depthMap = BuildDepthMap(flowchartData, layoutData);
            var horizontalOrder = BuildHorizontalOrder(layoutData, depthMap);
            var layers = BuildLayers(layoutData, depthMap, horizontalOrder);

            ApplyLayerPositions(layers, shapeMap, page);
        }

        private GraphLayoutData BuildGraphLayoutData(
            MermaidParser.FlowchartData flowchartData,
            Dictionary<string, Visio.Shape> shapeMap)
        {
            var layoutData = new GraphLayoutData();

            foreach (var entry in flowchartData.Nodes.Select((node, index) => new { node.Id, Index = index }))
            {
                layoutData.NodeOrder[entry.Id] = entry.Index;
            }

            foreach (var nodeId in shapeMap.Keys)
            {
                layoutData.ValidNodeIds.Add(nodeId);
                layoutData.IncomingEdges[nodeId] = new List<string>();
                layoutData.OutgoingEdges[nodeId] = new List<string>();
            }

            foreach (var connection in flowchartData.Connections)
            {
                if (!layoutData.ValidNodeIds.Contains(connection.FromId) ||
                    !layoutData.ValidNodeIds.Contains(connection.ToId) ||
                    string.Equals(connection.FromId, connection.ToId, StringComparison.Ordinal))
                {
                    continue;
                }

                layoutData.OutgoingEdges[connection.FromId].Add(connection.ToId);
                layoutData.IncomingEdges[connection.ToId].Add(connection.FromId);
            }

            return layoutData;
        }

        private Dictionary<string, int> BuildDepthMap(
            MermaidParser.FlowchartData flowchartData,
            GraphLayoutData layoutData)
        {
            var depthMap = layoutData.ValidNodeIds.ToDictionary(id => id, _ => -1, StringComparer.Ordinal);
            var roots = FindRootNodes(flowchartData, layoutData);
            var queue = new Queue<string>();

            foreach (var root in roots)
            {
                depthMap[root] = 0;
                queue.Enqueue(root);
            }

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                foreach (var next in layoutData.OutgoingEdges[current])
                {
                    if (depthMap[next] >= 0)
                    {
                        continue;
                    }

                    depthMap[next] = depthMap[current] + 1;
                    queue.Enqueue(next);
                }
            }

            foreach (var nodeId in layoutData.ValidNodeIds.Where(id => depthMap[id] < 0))
            {
                depthMap[nodeId] = 0;
            }

            return depthMap;
        }

        private List<string> FindRootNodes(
            MermaidParser.FlowchartData flowchartData,
            GraphLayoutData layoutData)
        {
            var roots = flowchartData.Nodes
                .Where(node => layoutData.ValidNodeIds.Contains(node.Id) && layoutData.IncomingEdges[node.Id].Count == 0)
                .Select(node => node.Id)
                .ToList();

            if (roots.Count > 0)
            {
                return roots;
            }

            return flowchartData.Nodes
                .Where(node => layoutData.ValidNodeIds.Contains(node.Id))
                .Select(node => node.Id)
                .Take(1)
                .ToList();
        }

        private Dictionary<string, double> BuildHorizontalOrder(
            GraphLayoutData layoutData,
            Dictionary<string, int> depthMap)
        {
            var orderedNodeIds = layoutData.ValidNodeIds
                .OrderBy(id => depthMap[id])
                .ThenBy(id => layoutData.NodeOrder[id])
                .ToList();

            var horizontalOrder = new Dictionary<string, double>(StringComparer.Ordinal);
            foreach (var nodeId in orderedNodeIds)
            {
                if (layoutData.IncomingEdges[nodeId].Count == 0)
                {
                    horizontalOrder[nodeId] = layoutData.NodeOrder[nodeId];
                    continue;
                }

                horizontalOrder[nodeId] = layoutData.IncomingEdges[nodeId]
                    .Where(horizontalOrder.ContainsKey)
                    .DefaultIfEmpty(nodeId)
                    .Average(parentId => horizontalOrder.ContainsKey(parentId)
                        ? horizontalOrder[parentId]
                        : layoutData.NodeOrder[nodeId]);
            }

            return horizontalOrder;
        }

        private List<List<string>> BuildLayers(
            GraphLayoutData layoutData,
            Dictionary<string, int> depthMap,
            Dictionary<string, double> horizontalOrder)
        {
            return layoutData.ValidNodeIds
                .OrderBy(id => depthMap[id])
                .ThenBy(id => layoutData.NodeOrder[id])
                .GroupBy(id => depthMap[id])
                .OrderBy(group => group.Key)
                .Select(group => group
                    .OrderBy(id => horizontalOrder[id])
                    .ThenBy(id => layoutData.NodeOrder[id])
                    .ToList())
                .ToList();
        }

        private void ApplyLayerPositions(
            List<List<string>> layers,
            Dictionary<string, Visio.Shape> shapeMap,
            Visio.Page page)
        {
            double pageCenterX = GetPageCenterX(page);
            double totalHeight = CalculateTotalHeight(layers, shapeMap);
            double currentY = totalHeight + LayoutTopPadding;

            foreach (var layer in layers)
            {
                double maxHeight = GetLayerMaxHeight(layer, shapeMap);
                double totalWidth = CalculateLayerWidth(layer, shapeMap);
                double currentX = pageCenterX - totalWidth / 2.0;

                foreach (var nodeId in layer)
                {
                    var shape = shapeMap[nodeId];
                    double width = shape.CellsU["Width"].ResultIU;

                    currentX += width / 2.0;
                    shape.CellsU["PinX"].ResultIU = currentX;
                    shape.CellsU["PinY"].ResultIU = currentY - maxHeight / 2.0;
                    currentX += width / 2.0 + HorizontalGap;
                }

                currentY -= maxHeight + VerticalGap;
            }
        }

        private double GetPageCenterX(Visio.Page page)
        {
            try
            {
                return page.PageSheet.CellsU["PageWidth"].ResultIU / 2.0;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"读取页面中心点失败，改用默认值: {ex.Message}");
                return DefaultPageCenterX;
            }
        }

        private double CalculateTotalHeight(List<List<string>> layers, Dictionary<string, Visio.Shape> shapeMap)
        {
            return layers.Sum(layer => GetLayerMaxHeight(layer, shapeMap))
                + VerticalGap * Math.Max(0, layers.Count - 1);
        }

        private double GetLayerMaxHeight(List<string> layer, Dictionary<string, Visio.Shape> shapeMap)
        {
            return layer.Max(id => shapeMap[id].CellsU["Height"].ResultIU);
        }

        private double CalculateLayerWidth(List<string> layer, Dictionary<string, Visio.Shape> shapeMap)
        {
            return layer.Sum(id => shapeMap[id].CellsU["Width"].ResultIU)
                + HorizontalGap * Math.Max(0, layer.Count - 1);
        }
    }
}
