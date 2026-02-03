using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace aribeth
{
    public sealed class TwoDAData
    {
        public List<string> Columns { get; } = new();
        public List<TwoDARow> Rows { get; } = new();
        public int HeaderIndent { get; set; } = 1;
        public int IndexWidth { get; set; }
        public List<int> ColumnWidths { get; } = new();
        public List<int> HeaderTokenStarts { get; } = new();
        public int HeaderVisualLength { get; set; }
    }

    public sealed class TwoDARow
    {
        public string Index { get; set; } = string.Empty;
        public List<string> Values { get; } = new();
        public List<int> TokenStarts { get; } = new();
        public int VisualLength { get; set; }
    }

    public static class TwoDAParser
    {
        public static TwoDAData Parse(string filePath, Action<string>? log = null, bool strictHeader = false)
        {
            var data = new TwoDAData();
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            var lineIndex = 0;

            string? nextNonEmpty()
            {
                while (lineIndex < lines.Length)
                {
                    var current = lines[lineIndex++];
                    if (string.IsNullOrWhiteSpace(current))
                    {
                        continue;
                    }

                    return current;
                }

                return null;
            }

            var header = nextNonEmpty();
            if (header == null || !header.TrimStart().StartsWith("2DA", StringComparison.OrdinalIgnoreCase))
            {
                if (strictHeader)
                {
                    throw new InvalidDataException("Missing 2DA header.");
                }

                log?.Invoke("Missing 2DA header. File may not be a valid 2DA.");
            }

            var columnsLine = nextNonEmpty();
            if (columnsLine == null)
            {
                log?.Invoke("No column header line found.");
                return data;
            }

            var headerTokens = TokenizeWithPositions(columnsLine);
            data.Columns.AddRange(headerTokens.Tokens.Select(token => token.Text));
            if (data.Columns.Count == 0)
            {
                log?.Invoke("Column header line is empty.");
                return data;
            }

            data.HeaderIndent = headerTokens.Tokens.Count > 0 ? headerTokens.Tokens[0].Start : 1;
            data.HeaderVisualLength = headerTokens.VisualLength;
            data.HeaderTokenStarts.Clear();
            data.HeaderTokenStarts.AddRange(headerTokens.Tokens.Select(token => token.Start));
            for (var i = 0; i < data.Columns.Count; i++)
            {
                data.ColumnWidths.Add(GetTokenWidth(headerTokens.Tokens, i));
            }

            while (lineIndex < lines.Length)
            {
                var rawLine = lines[lineIndex++];
                if (string.IsNullOrWhiteSpace(rawLine))
                {
                    continue;
                }

                var tokens = Tokenize(rawLine);
                if (tokens.Count == 0)
                {
                    continue;
                }

                var row = new TwoDARow { Index = tokens[0] };
                var remaining = tokens.Skip(1).ToList();

                if (remaining.Count < data.Columns.Count)
                {
                    var missing = data.Columns.Count - remaining.Count;
                    for (var i = 0; i < missing; i++)
                    {
                        remaining.Add("****");
                    }
                }
                else if (remaining.Count > data.Columns.Count)
                {
                    log?.Invoke($"Row {row.Index} has extra tokens; extra values will be concatenated.");
                    var combined = string.Join(" ", remaining.Skip(data.Columns.Count - 1));
                    remaining = remaining.Take(data.Columns.Count - 1).ToList();
                    remaining.Add(combined);
                }

                row.Values.AddRange(remaining);
                data.Rows.Add(row);

                var widthTokens = TokenizeWithPositions(rawLine);
                if (widthTokens.Tokens.Count > 0)
                {
                    data.IndexWidth = Math.Max(data.IndexWidth, GetTokenWidth(widthTokens.Tokens, 0));
                }

                for (var i = 0; i < data.Columns.Count; i++)
                {
                    if (data.ColumnWidths.Count <= i)
                    {
                        data.ColumnWidths.Add(0);
                    }

                    var tokenIndex = i + 1;
                    if (tokenIndex < widthTokens.Tokens.Count)
                    {
                        data.ColumnWidths[i] = Math.Max(data.ColumnWidths[i], GetTokenWidth(widthTokens.Tokens, tokenIndex));
                    }
                }

                row.TokenStarts.Clear();
                row.TokenStarts.AddRange(widthTokens.Tokens.Select(token => token.Start));
                row.VisualLength = widthTokens.VisualLength;
            }

            return data;
        }

        public static string Serialize(TwoDAData data)
        {
            var (indexWidth, columnWidths, headerIndent) = CalculateWidths(data);
            var useHeaderLayout = data.HeaderTokenStarts.Count == data.Columns.Count;
            var builder = new StringBuilder();
            builder.AppendLine("2DA V2.0");
            builder.AppendLine();

            builder.AppendLine(BuildLineFromStarts(
                data.Columns,
                data.HeaderTokenStarts,
                headerIndent,
                data.HeaderVisualLength,
                columnWidths,
                useColumnWidths: true));

            for (var i = 0; i < data.Rows.Count; i++)
            {
                var row = data.Rows[i];
                var index = string.IsNullOrWhiteSpace(row.Index) ? i.ToString() : row.Index;
                var tokens = new List<string> { index };
                for (var col = 0; col < data.Columns.Count; col++)
                {
                    tokens.Add(col < row.Values.Count ? row.Values[col] : "****");
                }

                var rowStarts = useHeaderLayout
                    ? new List<int> { 0 }.Concat(data.HeaderTokenStarts).ToList()
                    : row.TokenStarts;
                builder.AppendLine(BuildLineFromStarts(
                    tokens,
                    rowStarts,
                    0,
                    row.VisualLength,
                    columnWidths,
                    indexWidth,
                    useColumnWidths: true));
            }

            return builder.ToString();
        }

        private static List<string> Tokenize(string line)
        {
            var tokens = new List<string>();
            var builder = new StringBuilder();
            var inQuotes = false;

            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];

                if (c == '"')
                {
                    if (builder.Length == 0)
                    {
                        // Keep quote as part of the token.
                        builder.Append(c);
                    }
                    else
                    {
                        builder.Append(c);
                    }

                    inQuotes = !inQuotes;
                    continue;
                }

                if (!inQuotes && char.IsWhiteSpace(c))
                {
                    if (builder.Length > 0)
                    {
                        tokens.Add(builder.ToString());
                        builder.Clear();
                    }

                    continue;
                }

                builder.Append(c);
            }

            if (builder.Length > 0)
            {
                tokens.Add(builder.ToString());
            }

            return tokens;
        }

        private static TokenizeResult TokenizeWithPositions(string line)
        {
            var tokens = new List<Token>();
            var builder = new StringBuilder();
            var inQuotes = false;
            var tokenStart = -1;
            var visualIndex = 0;
            const int tabWidth = 8;

            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];

                if (c == '"')
                {
                    if (builder.Length == 0)
                    {
                        tokenStart = visualIndex;
                    }

                    builder.Append(c);
                    inQuotes = !inQuotes;
                    visualIndex++;
                    continue;
                }

                if (!inQuotes && char.IsWhiteSpace(c))
                {
                    if (builder.Length > 0)
                    {
                        tokens.Add(new Token(builder.ToString(), tokenStart));
                        builder.Clear();
                        tokenStart = -1;
                    }

                    if (c == '\t')
                    {
                        visualIndex = ((visualIndex / tabWidth) + 1) * tabWidth;
                    }
                    else
                    {
                        visualIndex++;
                    }
                    continue;
                }

                if (builder.Length == 0)
                {
                    tokenStart = visualIndex;
                }

                builder.Append(c);
                visualIndex++;
            }

            if (builder.Length > 0)
            {
                tokens.Add(new Token(builder.ToString(), tokenStart));
            }

            return new TokenizeResult(tokens, visualIndex);
        }

        private static int GetTokenWidth(List<Token> tokens, int index)
        {
            if (index < 0 || index >= tokens.Count)
            {
                return 0;
            }

            if (index + 1 < tokens.Count)
            {
                return Math.Max(tokens[index + 1].Start - tokens[index].Start, tokens[index].Text.Length);
            }

            return tokens[index].Text.Length;
        }

        private static (int IndexWidth, List<int> ColumnWidths, int HeaderIndent) CalculateWidths(TwoDAData data)
        {
            var headerIndent = Math.Max(0, data.HeaderIndent);
            var columnWidths = new List<int>();
            var columnCount = data.Columns.Count;
            for (var i = 0; i < columnCount; i++)
            {
                var nameLength = data.Columns[i]?.Length ?? 0;
                var maxLength = nameLength;
                foreach (var row in data.Rows)
                {
                    if (row.Values.Count > i)
                    {
                        maxLength = Math.Max(maxLength, row.Values[i]?.Length ?? 0);
                    }
                }

                var existing = i < data.ColumnWidths.Count ? data.ColumnWidths[i] : 0;
                var width = Math.Max(existing, maxLength);
                columnWidths.Add(width);
            }

            var maxIndex = 0;
            foreach (var row in data.Rows)
            {
                maxIndex = Math.Max(maxIndex, row.Index?.Length ?? 0);
            }

            var indexWidth = Math.Max(data.IndexWidth, maxIndex);

            if (data.HeaderTokenStarts.Count == columnCount && columnCount > 0)
            {
                columnWidths.Clear();
                for (var i = 0; i < columnCount; i++)
                {
                    var start = data.HeaderTokenStarts[i];
                    var nextStart = i + 1 < columnCount ? data.HeaderTokenStarts[i + 1] : data.HeaderVisualLength;
                    var headerWidth = Math.Max(0, nextStart - start);
                    var nameLength = data.Columns[i]?.Length ?? 0;
                    var maxLength = nameLength;
                    foreach (var row in data.Rows)
                    {
                        if (row.Values.Count > i)
                        {
                            maxLength = Math.Max(maxLength, row.Values[i]?.Length ?? 0);
                        }
                    }

                    var width = Math.Max(headerWidth, maxLength);
                    columnWidths.Add(width);
                }

                indexWidth = Math.Max(indexWidth, headerIndent);
            }

            return (indexWidth, columnWidths, headerIndent);
        }

        private static string BuildLineFromStarts(
            List<string> tokens,
            List<int> tokenStarts,
            int fallbackIndent,
            int fallbackVisualLength,
            List<int> columnWidths,
            int indexWidthOverride = 0,
            bool useColumnWidths = true)
        {
            var builder = new StringBuilder();
            var count = tokens.Count;
            var starts = new List<int>();
            var minStart = fallbackIndent;

            for (var i = 0; i < count; i++)
            {
                var desired = tokenStarts.Count > i ? tokenStarts[i] : minStart;
                var start = Math.Max(desired, minStart);
                starts.Add(start);
                var token = tokens[i] ?? string.Empty;
                minStart = start + token.Length;

                if (useColumnWidths)
                {
                    if (i == 0 && indexWidthOverride > 0)
                    {
                        minStart = Math.Max(minStart, indexWidthOverride);
                    }
                    else if (i - 1 >= 0 && i - 1 < columnWidths.Count)
                    {
                        minStart = Math.Max(minStart, starts[i - 1] + columnWidths[i - 1]);
                    }

                    if (i < count - 1)
                    {
                        minStart = Math.Max(minStart, start + token.Length + 1);
                    }
                }
            }

            for (var i = 0; i < count; i++)
            {
                var token = tokens[i] ?? string.Empty;
                var start = starts[i];
                if (builder.Length < start)
                {
                    builder.Append(' ', start - builder.Length);
                }

                builder.Append(token);
            }

            var targetLength = fallbackVisualLength > 0 ? fallbackVisualLength : builder.Length;
            if (builder.Length < targetLength)
            {
                builder.Append(' ', targetLength - builder.Length);
            }

            return builder.ToString();
        }

        private readonly struct TokenizeResult
        {
            public TokenizeResult(List<Token> tokens, int visualLength)
            {
                Tokens = tokens;
                VisualLength = visualLength;
            }

            public List<Token> Tokens { get; }
            public int VisualLength { get; }
        }

        private readonly struct Token
        {
            public Token(string text, int start)
            {
                Text = text;
                Start = start;
            }

            public string Text { get; }
            public int Start { get; }
        }
    }
}
