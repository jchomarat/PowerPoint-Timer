using System;
using System.Drawing;

namespace PowerPointTimer.Core
{
    public class Constants
    {
        public static string DefaultTimerValue = "04:20";
        
        public static string TimerGroup = "TimerGroup";
        public static string TimerGroupValue = "true";

        public static string TimerCounter = "TimerCounter";
        public static string TimerCounterValue = "true";
       
        public static string TimerLauncher = "TimerLauncher";
        public static string TimerLauncherValue = "true";

        public static TimeSpan RefreshTimeSpan = TimeSpan.FromSeconds(1);

        public static Color Less25Percent = Color.Orange;
        public static Color Less10Percent = Color.Red;
    }
}
