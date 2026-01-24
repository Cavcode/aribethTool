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
    }

    public sealed class TwoDARow
    {
        public string Index { get; set; } = string.Empty;
        public List<string> Values { get; } = new();
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
                    var current = lines[lineIndex++].Trim();
                    if (current.Length == 0)
                    {
                        continue;
                    }

                    return current;
                }

                return null;
            }

            var header = nextNonEmpty();
            if (header == null || !header.StartsWith("2DA", StringComparison.OrdinalIgnoreCase))
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

            data.Columns.AddRange(Tokenize(columnsLine));
            if (data.Columns.Count == 0)
            {
                log?.Invoke("Column header line is empty.");
                return data;
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
            }

            return data;
        }

        public static string Serialize(TwoDAData data)
        {
            var builder = new StringBuilder();
            builder.AppendLine("2DA V2.0");
            builder.AppendLine();

            var header = " " + string.Join("\t", data.Columns);
            builder.AppendLine(header);

            for (var i = 0; i < data.Rows.Count; i++)
            {
                var row = data.Rows[i];
                var index = string.IsNullOrWhiteSpace(row.Index) ? i.ToString() : row.Index;
                builder.Append(index);
                builder.Append('\t');
                builder.AppendLine(string.Join("\t", row.Values));
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
    }
}
