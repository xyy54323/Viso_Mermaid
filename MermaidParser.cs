using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VisioAddIn1
{
    public class MermaidParser
    {
        private const string DefaultDirection = "TD";
        private const string DefaultNodeShape = "rectangle";
        private const string InvalidNodeId = "-";
        private const int NodeIdGroupIndex = 1;
        private const int FirstNodeTextGroupIndex = 2;
        private const int LastNodeTextGroupIndex = 6;

        private const string NodeIdPattern = @"[A-Za-z0-9_-]+";
        private const string NodeShapePattern =
            @"(?:\[([^\]]*)\]|\{([^\}]*)\}|\(([^\)]*)\)|>([^<]*)<|(?:\[\[([^\]]*)\]\]))";

        public class Node
        {
            public string Id { get; set; }
            public string Text { get; set; }
            public string Shape { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
        }

        public class Connection
        {
            public string FromId { get; set; }
            public string ToId { get; set; }
            public string Label { get; set; }
        }

        public class FlowchartData
        {
            public List<Node> Nodes { get; set; } = new List<Node>();
            public List<Connection> Connections { get; set; } = new List<Connection>();
            public string Direction { get; set; } = DefaultDirection;
        }

        private sealed class ParseState
        {
            private readonly Dictionary<string, string> _nodeTexts = new Dictionary<string, string>(StringComparer.Ordinal);
            private readonly Dictionary<string, string> _nodeShapes = new Dictionary<string, string>(StringComparer.Ordinal);
            private readonly HashSet<string> _seenConnections = new HashSet<string>(StringComparer.Ordinal);

            public FlowchartData FlowchartData { get; } = new FlowchartData();

            public void UpsertNode(string id, string text, string shape)
            {
                if (string.IsNullOrWhiteSpace(id) || string.Equals(id, InvalidNodeId, StringComparison.Ordinal))
                {
                    return;
                }

                _nodeTexts[id] = string.IsNullOrWhiteSpace(text) ? id : text;
                _nodeShapes[id] = string.IsNullOrWhiteSpace(shape) ? DefaultNodeShape : shape;
            }

            public void EnsureDefaultNode(string id)
            {
                if (!_nodeTexts.ContainsKey(id))
                {
                    _nodeTexts[id] = id;
                    _nodeShapes[id] = DefaultNodeShape;
                }
            }

            public bool TryAddConnection(string fromId, string toId, string label)
            {
                string connectionKey = $"{fromId}->{toId}|{label}";
                return _seenConnections.Add(connectionKey);
            }

            public void FinalizeNodes()
            {
                foreach (var nodeId in _nodeTexts.Keys.Where(IsValidNodeId))
                {
                    FlowchartData.Nodes.Add(new Node
                    {
                        Id = nodeId,
                        Text = _nodeTexts[nodeId],
                        Shape = _nodeShapes[nodeId]
                    });
                }
            }

            private static bool IsValidNodeId(string nodeId)
            {
                return !string.IsNullOrWhiteSpace(nodeId) &&
                       !string.Equals(nodeId, InvalidNodeId, StringComparison.Ordinal);
            }
        }

        private static readonly Regex GraphHeaderRegex = new Regex(
            @"^(graph|flowchart)\s+([A-Za-z]{2})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex NodeDefinitionRegex = new Regex(
            $@"^\s*({NodeIdPattern})\s*{NodeShapePattern}\s*$",
            RegexOptions.Compiled);

        private static readonly Regex InlineNodeDefinitionRegex = new Regex(
            $@"({NodeIdPattern}){NodeShapePattern}",
            RegexOptions.Compiled);

        private static readonly Regex ConnectionRegex = new Regex(
            $@"({NodeIdPattern})\s*(?:==>|-->|->|--(?:-|>))\s*(?:\|([^|]*)\|\s*)?({NodeIdPattern})",
            RegexOptions.Compiled);

        public FlowchartData Parse(string mermaidCode)
        {
            var state = new ParseState();
            var lines = GetMeaningfulLines(mermaidCode).ToList();
            if (lines.Count == 0)
            {
                return state.FlowchartData;
            }

            int startIndex = TryParseHeader(lines[0], state.FlowchartData) ? 1 : 0;
            for (int i = startIndex; i < lines.Count; i++)
            {
                RegisterStandaloneNode(lines[i], state);
            }

            foreach (var line in lines.Skip(startIndex))
            {
                RegisterInlineNodes(line, state);
                RegisterConnections(line, state);
            }

            state.FinalizeNodes();
            return state.FlowchartData;
        }

        private IEnumerable<string> GetMeaningfulLines(string mermaidCode)
        {
            if (string.IsNullOrWhiteSpace(mermaidCode))
            {
                yield break;
            }

            foreach (var rawLine in mermaidCode.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith("%%", StringComparison.Ordinal))
                {
                    continue;
                }

                yield return line;
            }
        }

        private bool TryParseHeader(string line, FlowchartData result)
        {
            var match = GraphHeaderRegex.Match(line);
            if (!match.Success)
            {
                return false;
            }

            result.Direction = match.Groups[2].Value.ToUpperInvariant();
            return true;
        }

        private void RegisterStandaloneNode(string line, ParseState state)
        {
            var match = NodeDefinitionRegex.Match(line);
            if (!match.Success)
            {
                return;
            }

            RegisterNodeMatch(match, state);
        }

        private void RegisterInlineNodes(string line, ParseState state)
        {
            foreach (Match match in InlineNodeDefinitionRegex.Matches(line))
            {
                RegisterNodeMatch(match, state);
            }
        }

        private void RegisterNodeMatch(Match match, ParseState state)
        {
            state.UpsertNode(
                match.Groups[NodeIdGroupIndex].Value,
                ExtractNodeText(match),
                DetectShape(match.Value));
        }

        private void RegisterConnections(string line, ParseState state)
        {
            string normalizedLine = NormalizeConnectionLine(line);
            foreach (Match match in ConnectionRegex.Matches(normalizedLine))
            {
                string fromId = match.Groups[1].Value;
                string toId = match.Groups[3].Value;
                if (!IsValidConnectionEndpoint(fromId) || !IsValidConnectionEndpoint(toId))
                {
                    continue;
                }

                state.EnsureDefaultNode(fromId);
                state.EnsureDefaultNode(toId);

                string label = NormalizeConnectionLabel(match.Groups[2].Value);
                if (!state.TryAddConnection(fromId, toId, label))
                {
                    continue;
                }

                state.FlowchartData.Connections.Add(new Connection
                {
                    FromId = fromId,
                    ToId = toId,
                    Label = label
                });
            }
        }

        private bool IsValidConnectionEndpoint(string nodeId)
        {
            return !string.IsNullOrWhiteSpace(nodeId) &&
                   !string.Equals(nodeId, InvalidNodeId, StringComparison.Ordinal);
        }

        private string NormalizeConnectionLabel(string label)
        {
            return string.IsNullOrWhiteSpace(label) ? string.Empty : label.Trim();
        }

        private string ExtractNodeText(Match match)
        {
            for (int groupIndex = FirstNodeTextGroupIndex; groupIndex <= LastNodeTextGroupIndex; groupIndex++)
            {
                string groupValue = match.Groups[groupIndex].Value;
                if (!string.IsNullOrWhiteSpace(groupValue))
                {
                    return groupValue.Trim();
                }
            }

            return match.Groups[NodeIdGroupIndex].Value;
        }

        private string DetectShape(string token)
        {
            if (token.Contains("[[") && token.Contains("]]"))
            {
                return "database";
            }

            if (token.Contains("{") && token.Contains("}"))
            {
                return "diamond";
            }

            if (token.Contains("(") && token.Contains(")"))
            {
                return "rounded rectangle";
            }

            if (token.Contains(">") && token.Contains("<"))
            {
                return "circle";
            }

            return DefaultNodeShape;
        }

        private string NormalizeConnectionLine(string line)
        {
            return InlineNodeDefinitionRegex.Replace(line, "$1");
        }
    }
}
