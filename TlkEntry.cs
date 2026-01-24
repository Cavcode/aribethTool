using System;

namespace aribeth
{
    public sealed class TlkEntry
    {
        public int Index { get; set; }
        public int Length { get; set; }
        public string Text { get; set; } = string.Empty;
        public string SoundResRef { get; set; } = string.Empty;
        public string Duration { get; set; } = string.Empty;

        public void RecalculateLength()
        {
            Length = Text?.Length ?? 0;
        }
    }
}
