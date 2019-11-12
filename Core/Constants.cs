using System;
using System.Drawing;

namespace PowerPointTimer.Core
{
    public class Constants
    {
        public static string DefaultTimerValue = "04:20";
        public static string TimerTag = "TimerTag";
        public static string TimerTagValue = "TimerDigital";
        public static string TimerLauncherTag = "TimerLauncherTag";
        public static TimeSpan RefreshTimeSpan = TimeSpan.FromSeconds(1);

        public static Color Less25Percent = Color.Orange;
        public static Color Less10Percent = Color.Red;
    }
}
