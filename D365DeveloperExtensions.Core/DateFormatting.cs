using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace D365DeveloperExtensions.Core
{
    public static class DateFormatting
    {
        public static string MsToReadableTime(int ms)
        {
            TimeSpan ts = TimeSpan.FromMilliseconds(ms);

            var parts = $"{ts.Days:D2}d:{ts.Hours:D2}h:{ts.Minutes:D2}m:{ts.Seconds:D2}s:{ts.Milliseconds:D3}ms"
                .Split(':')
                .SkipWhile(s => Regex.Match(s, @"00\w").Success)
                .ToArray();
            var result = string.Join(" ", parts);

            return result;
        }
    }
}