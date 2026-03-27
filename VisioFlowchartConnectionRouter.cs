using System;
using System.Collections.Generic;
using System.Linq;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddIn1
{
    internal sealed class VisioFlowchartConnectionRouter
    {
        private const string LeftSide = "left";
        private const string RightSide = "right";
        private const string TopSide = "top";
        private const string BottomSide = "bottom";
        private const string DiamondShape = "diamond";

        private const string BlackThemeFormula = "THEMEGUARD(RGB(0,0,0))";
        private const string NoRoundingFormula = "0 in";
        private const string LinePatternFormula = "1";
        private const string LineWeightFormula = "1.5 pt";
        private const string ConnectorRouteExtensionFormula = "1";
        private const string EndArrowFormula = "13";
        private const string EndArrowSizeFormula = "3";
        private const string ConnectorTextSizeFormula = "11 pt";
        private const string NoArrowFormula = "0";
        private const string NoFillFormula = "0";
        private const string VerticalRouteStyle = "22";
        private const string HorizontalRouteStyle = "21";

        private readonly Visio.Application _application;

        private sealed class ConnectionBuckets
        {
            public List<MermaidParser.Connection> NormalConnections { get; } = new List<MermaidParser.Connection>();
            public List<MermaidParser.Connection> SelfLoopConnections { get; } = new List<MermaidParser.Connection>();
        }

        public VisioFlowchartConnectionRouter(Visio.Application application)
        {
            _application = application;
        }

        public void CreateConnections(
            Visio.Page page,
            MermaidParser.FlowchartData flowchartData,
            Dictionary<string, Visio.Shape> shapeMap)
        {
            var nodeTypeMap = flowchartData.Nodes.ToDictionary(node => node.Id, node => node.Shape, StringComparer.Ordinal);
            var connectionBuckets = BucketConnections(flowchartData.Connections, shapeMap);
            var outgoingCountByNode = BuildOutgoingCountMap(connectionBuckets.NormalConnections);
            var sideAssignments = AssignConnectionSides(
                connectionBuckets.NormalConnections,
                shapeMap,
                nodeTypeMap,
                outgoingCountByNode);

            DrawSelfLoops(page, connectionBuckets.SelfLoopConnections, shapeMap);
            DrawNormalConnections(page, connectionBuckets.NormalConnections, shapeMap, sideAssignments);
        }

        private ConnectionBuckets BucketConnections(
            IEnumerable<MermaidParser.Connection> connections,
            Dictionary<string, Visio.Shape> shapeMap)
        {
            var buckets = new ConnectionBuckets();

            foreach (var connection in connections)
            {
                if (!shapeMap.ContainsKey(connection.FromId) || !shapeMap.ContainsKey(connection.ToId))
                {
                    continue;
                }

                if (string.Equals(connection.FromId, connection.ToId, StringComparison.Ordinal))
                {
                    buckets.SelfLoopConnections.Add(connection);
                }
                else
                {
                    buckets.NormalConnections.Add(connection);
                }
            }

            return buckets;
        }

        private Dictionary<string, int> BuildOutgoingCountMap(IEnumerable<MermaidParser.Connection> connections)
        {
            return connections
                .GroupBy(connection => connection.FromId)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        }

        private void DrawSelfLoops(
            Visio.Page page,
            IEnumerable<MermaidParser.Connection> selfLoopConnections,
            Dictionary<string, Visio.Shape> shapeMap)
        {
            foreach (var connection in selfLoopConnections)
            {
                CreateSelfLoopConnector(page, shapeMap[connection.FromId], connection.Label);
            }
        }

        private void DrawNormalConnections(
            Visio.Page page,
            IEnumerable<MermaidParser.Connection> normalConnections,
            Dictionary<string, Visio.Shape> shapeMap,
            Dictionary<MermaidParser.Connection, (string startSide, string endSide)> sideAssignments)
        {
            foreach (var connection in normalConnections)
            {
                var assignment = sideAssignments[connection];
                CreateConnector(
                    page,
                    shapeMap[connection.FromId],
                    shapeMap[connection.ToId],
                    connection.Label,
                    assignment.startSide,
                    assignment.endSide);
            }
        }

        private void CreateConnector(
            Visio.Page page,
            Visio.Shape fromShape,
            Visio.Shape toShape,
            string label,
            string startSide,
            string endSide)
        {
            try
            {
                var connector = page.Drop(_application.ConnectorToolDataObject, 0, 0);
                var startGlue = GetGluePoint(startSide);
                var endGlue = GetGluePoint(endSide);

                connector.CellsU["BeginX"].GlueToPos(fromShape, startGlue.x, startGlue.y);
                connector.CellsU["EndX"].GlueToPos(toShape, endGlue.x, endGlue.y);
                connector.CellsU["Rounding"].Formula = NoRoundingFormula;
                connector.CellsU["ShapeRouteStyle"].Formula = GetConnectorRouteStyle(fromShape, toShape);

                if (!string.IsNullOrEmpty(label))
                {
                    connector.Text = label;
                }

                SetConnectorStyle(connector);
            }
            catch (Exception ex)
            {
                InternalLog.Info($"创建连接线失败: {ex.Message}");
            }
        }

        private void CreateSelfLoopConnector(Visio.Page page, Visio.Shape shape, string label)
        {
            try
            {
                double centerX = shape.CellsU["PinX"].ResultIU;
                double centerY = shape.CellsU["PinY"].ResultIU;
                double halfWidth = shape.CellsU["Width"].ResultIU / 2.0;
                double halfHeight = shape.CellsU["Height"].ResultIU / 2.0;
                double offsetX = Math.Max(0.6, halfWidth * 0.9);
                double offsetY = Math.Max(0.45, halfHeight * 0.8);
                double[] xyArray =
                {
                    centerX + halfWidth, centerY,
                    centerX + halfWidth + offsetX, centerY,
                    centerX + halfWidth + offsetX, centerY + halfHeight + offsetY,
                    centerX, centerY + halfHeight + offsetY,
                    centerX, centerY + halfHeight
                };

                var connector = page.DrawPolyline(xyArray, 0);

                if (!string.IsNullOrWhiteSpace(label))
                {
                    connector.Text = label;
                    connector.CellsU["TxtPinX"].FormulaU = "Width*0.60";
                    connector.CellsU["TxtPinY"].FormulaU = "Height*1.12";
                    connector.CellsU["TxtLocPinX"].FormulaU = "TxtWidth*0.5";
                    connector.CellsU["TxtLocPinY"].FormulaU = "TxtHeight*0.5";
                }

                SetConnectorStyle(connector);
                connector.CellsU["BeginArrow"].Formula = NoArrowFormula;
                connector.CellsU["LineColor"].FormulaForceU = BlackThemeFormula;
                connector.CellsU["Char.Color"].FormulaForceU = BlackThemeFormula;
                connector.CellsU["FillPattern"].Formula = NoFillFormula;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"创建回环线失败: {ex.Message}");
            }
        }

        private void SetConnectorStyle(Visio.Shape connector)
        {
            try
            {
                connector.CellsU["LineColor"].FormulaForceU = BlackThemeFormula;
                connector.CellsU["LinePattern"].Formula = LinePatternFormula;
                connector.CellsU["LineWeight"].Formula = LineWeightFormula;
                connector.CellsU["ConLineRouteExt"].Formula = ConnectorRouteExtensionFormula;
                connector.CellsU["Rounding"].Formula = NoRoundingFormula;
                connector.CellsU["EndArrow"].Formula = EndArrowFormula;
                connector.CellsU["EndArrowSize"].Formula = EndArrowSizeFormula;
                connector.CellsU["Char.Color"].FormulaForceU = BlackThemeFormula;
                connector.CellsU["Char.Style"].Formula = "0";
                connector.CellsU["Char.Size"].Formula = ConnectorTextSizeFormula;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"设置连接线样式失败: {ex.Message}");
            }
        }

        private string GetConnectorRouteStyle(Visio.Shape fromShape, Visio.Shape toShape)
        {
            try
            {
                double deltaX = Math.Abs(fromShape.CellsU["PinX"].ResultIU - toShape.CellsU["PinX"].ResultIU);
                double deltaY = Math.Abs(fromShape.CellsU["PinY"].ResultIU - toShape.CellsU["PinY"].ResultIU);
                return deltaY >= deltaX ? VerticalRouteStyle : HorizontalRouteStyle;
            }
            catch (Exception ex)
            {
                InternalLog.Info($"计算连接线路由样式失败，改用默认垂直路由: {ex.Message}");
                return VerticalRouteStyle;
            }
        }

        private Dictionary<MermaidParser.Connection, (string startSide, string endSide)> AssignConnectionSides(
            List<MermaidParser.Connection> connections,
            Dictionary<string, Visio.Shape> shapeMap,
            Dictionary<string, string> nodeTypeMap,
            Dictionary<string, int> outgoingCountByNode)
        {
            var assignments = new Dictionary<MermaidParser.Connection, (string startSide, string endSide)>();
            var usedSides = shapeMap.Keys.ToDictionary(id => id, _ => new HashSet<string>(), StringComparer.Ordinal);
            var reservedTopIncoming = ReservePreferredTopIncomingConnections(connections, shapeMap, nodeTypeMap);
            var orderedConnections = OrderConnectionsForAssignment(connections, shapeMap, reservedTopIncoming);

            foreach (var connection in orderedConnections)
            {
                var fromShape = shapeMap[connection.FromId];
                var toShape = shapeMap[connection.ToId];
                string fromShapeType = nodeTypeMap[connection.FromId];
                string toShapeType = nodeTypeMap[connection.ToId];
                int outgoingCount = outgoingCountByNode.TryGetValue(connection.FromId, out var count) ? count : 0;

                var startPreferences = GetSidePreferences(fromShape, toShape, fromShapeType, isStart: true, outgoingCount);
                var endPreferences = GetSidePreferences(fromShape, toShape, toShapeType, isStart: false);

                string startSide = ChooseAvailableSide(startPreferences, usedSides[connection.FromId]);
                string endSide = ChooseAvailableSide(endPreferences, usedSides[connection.ToId]);

                usedSides[connection.FromId].Add(startSide);
                usedSides[connection.ToId].Add(endSide);
                assignments[connection] = (startSide, endSide);
            }

            return assignments;
        }

        private HashSet<MermaidParser.Connection> ReservePreferredTopIncomingConnections(
            IEnumerable<MermaidParser.Connection> connections,
            Dictionary<string, Visio.Shape> shapeMap,
            Dictionary<string, string> nodeTypeMap)
        {
            var reservedConnections = new HashSet<MermaidParser.Connection>();

            foreach (var group in connections.GroupBy(connection => connection.ToId))
            {
                if (IsDiamond(nodeTypeMap[group.Key]))
                {
                    continue;
                }

                var preferredTopConnection = group
                    .OrderBy(connection => Math.Abs(
                        shapeMap[connection.FromId].CellsU["PinX"].ResultIU -
                        shapeMap[connection.ToId].CellsU["PinX"].ResultIU))
                    .ThenByDescending(connection => IsMostlyVertical(shapeMap[connection.FromId], shapeMap[connection.ToId]))
                    .FirstOrDefault();

                if (preferredTopConnection != null)
                {
                    reservedConnections.Add(preferredTopConnection);
                }
            }

            return reservedConnections;
        }

        private List<MermaidParser.Connection> OrderConnectionsForAssignment(
            IEnumerable<MermaidParser.Connection> connections,
            Dictionary<string, Visio.Shape> shapeMap,
            HashSet<MermaidParser.Connection> reservedTopIncoming)
        {
            return connections
                .OrderByDescending(connection => reservedTopIncoming.Contains(connection))
                .ThenByDescending(connection => IsMostlyVertical(shapeMap[connection.FromId], shapeMap[connection.ToId]))
                .ThenBy(connection => shapeMap[connection.FromId].CellsU["PinY"].ResultIU)
                .ThenBy(connection => shapeMap[connection.FromId].CellsU["PinX"].ResultIU)
                .ToList();
        }

        private bool IsMostlyVertical(Visio.Shape fromShape, Visio.Shape toShape)
        {
            double fromX = fromShape.CellsU["PinX"].ResultIU;
            double fromY = fromShape.CellsU["PinY"].ResultIU;
            double toX = toShape.CellsU["PinX"].ResultIU;
            double toY = toShape.CellsU["PinY"].ResultIU;
            return Math.Abs(toY - fromY) >= Math.Abs(toX - fromX);
        }

        private List<string> GetSidePreferences(
            Visio.Shape fromShape,
            Visio.Shape toShape,
            string shapeType,
            bool isStart,
            int outgoingCount = 0)
        {
            var (primaryHorizontalSide, secondaryHorizontalSide) = GetHorizontalPreferences(fromShape, toShape, isStart);

            if (IsDiamond(shapeType))
            {
                return isStart
                    ? GetDiamondStartPreferences(fromShape, toShape, primaryHorizontalSide, secondaryHorizontalSide, outgoingCount)
                    : GetDiamondEndPreferences(toShape, fromShape, primaryHorizontalSide, secondaryHorizontalSide);
            }

            return isStart
                ? BuildPreferenceOrder(
                    fromShape,
                    toShape,
                    new[] { BottomSide },
                    new[] { primaryHorizontalSide, secondaryHorizontalSide, TopSide })
                : BuildPreferenceOrder(
                    toShape,
                    fromShape,
                    new[] { TopSide },
                    new[] { primaryHorizontalSide, secondaryHorizontalSide, BottomSide });
        }

        private bool IsDiamond(string shapeType)
        {
            return string.Equals(shapeType, DiamondShape, StringComparison.OrdinalIgnoreCase);
        }

        private List<string> GetDiamondStartPreferences(
            Visio.Shape fromShape,
            Visio.Shape toShape,
            string primaryHorizontalSide,
            string secondaryHorizontalSide,
            int outgoingCount)
        {
            if (outgoingCount <= 1)
            {
                return BuildPreferenceOrder(
                    fromShape,
                    toShape,
                    new[] { BottomSide },
                    new[] { primaryHorizontalSide, secondaryHorizontalSide, TopSide });
            }

            return BuildPreferenceOrder(
                fromShape,
                toShape,
                new[] { primaryHorizontalSide, secondaryHorizontalSide },
                new[] { BottomSide, TopSide });
        }

        private List<string> GetDiamondEndPreferences(
            Visio.Shape currentShape,
            Visio.Shape otherShape,
            string primaryHorizontalSide,
            string secondaryHorizontalSide)
        {
            return BuildPreferenceOrder(
                currentShape,
                otherShape,
                new[] { TopSide },
                new[] { primaryHorizontalSide, secondaryHorizontalSide, BottomSide });
        }

        private (string primary, string secondary) GetHorizontalPreferences(
            Visio.Shape fromShape,
            Visio.Shape toShape,
            bool isStart)
        {
            double deltaX = toShape.CellsU["PinX"].ResultIU - fromShape.CellsU["PinX"].ResultIU;

            if (deltaX < 0)
            {
                return isStart ? (LeftSide, RightSide) : (RightSide, LeftSide);
            }

            return isStart ? (RightSide, LeftSide) : (LeftSide, RightSide);
        }

        private List<string> BuildPreferenceOrder(
            Visio.Shape currentShape,
            Visio.Shape otherShape,
            IEnumerable<string> primarySides,
            IEnumerable<string> fallbackSides)
        {
            var orderedSides = new List<string>();

            foreach (var side in primarySides.Where(IsSpecifiedSide))
            {
                if (!orderedSides.Contains(side))
                {
                    orderedSides.Add(side);
                }
            }

            foreach (var side in fallbackSides
                .Where(side => IsSpecifiedSide(side) && !orderedSides.Contains(side))
                .OrderBy(side => GetSideDistance(currentShape, otherShape, side)))
            {
                orderedSides.Add(side);
            }

            return orderedSides;
        }

        private bool IsSpecifiedSide(string side)
        {
            return !string.IsNullOrWhiteSpace(side);
        }

        private double GetSideDistance(Visio.Shape currentShape, Visio.Shape otherShape, string side)
        {
            var point = GetSidePoint(currentShape, side);
            double otherX = otherShape.CellsU["PinX"].ResultIU;
            double otherY = otherShape.CellsU["PinY"].ResultIU;
            double deltaX = point.x - otherX;
            double deltaY = point.y - otherY;
            return deltaX * deltaX + deltaY * deltaY;
        }

        private (double x, double y) GetSidePoint(Visio.Shape shape, string side)
        {
            double centerX = shape.CellsU["PinX"].ResultIU;
            double centerY = shape.CellsU["PinY"].ResultIU;
            double halfWidth = shape.CellsU["Width"].ResultIU / 2.0;
            double halfHeight = shape.CellsU["Height"].ResultIU / 2.0;

            switch (side)
            {
                case LeftSide:
                    return (centerX - halfWidth, centerY);
                case RightSide:
                    return (centerX + halfWidth, centerY);
                case BottomSide:
                    return (centerX, centerY - halfHeight);
                default:
                    return (centerX, centerY + halfHeight);
            }
        }

        private string ChooseAvailableSide(List<string> preferences, HashSet<string> usedSides)
        {
            foreach (var side in preferences)
            {
                if (!usedSides.Contains(side))
                {
                    return side;
                }
            }

            return preferences[0];
        }

        private (double x, double y) GetGluePoint(string side)
        {
            switch (side)
            {
                case LeftSide:
                    return (0.0, 0.5);
                case RightSide:
                    return (1.0, 0.5);
                case BottomSide:
                    return (0.5, 0.0);
                default:
                    return (0.5, 1.0);
            }
        }
    }
}
