using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CrmDeveloperExtensions2.Core
{
    public static class DateFormatting
    {
        public static string MsToReadableTime(int ms)
        {
            TimeSpan ts = TimeSpan.FromMilliseconds(ms);

            var parts = string
                .Format("{0:D2}d:{1:D2}h:{2:D2}m:{3:D2}s:{4:D3}ms",
                    ts.Days, ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds)
                .Split(':')
                .SkipWhile(s => Regex.Match(s, @"00\w").Success)
                .ToArray();
            var result = string.Join(" ", parts);

            return result;
        }
    }
}