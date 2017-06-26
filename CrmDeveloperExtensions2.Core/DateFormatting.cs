using System;

namespace CrmDeveloperExtensions2.Core
{
    public static class DateFormatting
    {
        public static string MsToReadableTime(int ms)
        {
            TimeSpan time = TimeSpan.FromMilliseconds(ms);

            string result = time.ToString(@"hh\:mm\:ss\:fff");

            return result;
        }
    }
}