using System;
using System.Linq;
using System.Drawing;
using System.Globalization;
using System.Security.Policy;
using System.Timers;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PowerPointTimer.Core
{
    public enum CountDownStatusEnum
    {
        Invalid, Initialized, Ready, Running
    }

    public class CountDown
    {
        string timerDuration = String.Empty;
        TimeSpan timerDurationTimeSpan;
        Timer timer;
        Shape timerShape;
        Shape launcherShape;

        readonly string empty = " ";
        readonly string noTimeLeft = "00:00";

        public CountDown(Shape TimerShape)
        {
            timerShape = TimerShape;
        }

        public int UnderlyingShapeId { get { return timerShape.Id; } }

        public CountDownStatusEnum Status { get; set; } = CountDownStatusEnum.Initialized;

        public void Start()
        {
            if (Status == CountDownStatusEnum.Ready)
            {
                timerShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                if (launcherShape != null)
                    launcherShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Start timer and ignite Count down
                timer = new Timer(Constants.RefreshTimeSpan.TotalMilliseconds);
                timer.Elapsed += OnTimerElapsed;
                timer.Start();
                Status = CountDownStatusEnum.Running;
            }
        }

        public void Init(Shape LauncherShape)
        {
            launcherShape = LauncherShape;

            if (launcherShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                // Copy duration from launcher to actual countdown
                timerShape.TextFrame.TextRange.Text = launcherShape.TextFrame.TextRange.Text;

                // Get duration
                timerDuration = timerShape.TextFrame.TextRange.Text;

                // Convert it to TimeSpan
                if (!TimeSpan.TryParseExact(timerDuration, "mm\\:ss", CultureInfo.InvariantCulture, out timerDurationTimeSpan))
                {
                    Status = CountDownStatusEnum.Invalid;
                }
                else
                {
                    launcherShape = LauncherShape;
                    // copy style & position from the launcher to the actual countdown
                    timerShape.Width = launcherShape.Width;
                    timerShape.Height = launcherShape.Height;
                    timerShape.Left = launcherShape.Left;
                    timerShape.Top = launcherShape.Top;

                    launcherShape.PickUp();
                    timerShape.Apply();

                    Status = CountDownStatusEnum.Ready;
                }
            }
            else Status = CountDownStatusEnum.Invalid;
        }

        public void Stop()
        {
            timerShape.TextFrame.TextRange.Text = timerDuration;

            // Restore launcher view
            timerShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            if (launcherShape != null)
                launcherShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;

            if (timer != null) timer.Dispose();
            Status = CountDownStatusEnum.Ready;
        }

        void OnTimerElapsed(object o, ElapsedEventArgs args)
        {
            string currentTimeValue = timerShape.TextFrame.TextRange.Text;
            if (currentTimeValue == empty) currentTimeValue = noTimeLeft;

            if (TimeSpan.TryParseExact(currentTimeValue, "mm\\:ss", CultureInfo.InvariantCulture, out TimeSpan timeLeft))
            {
                if (timeLeft.TotalMilliseconds == 0)
                {
                    // We are done - we keep the timer however to make the countdown flicker
                    timerShape.TextFrame.TextRange.Text =
                        (timerShape.TextFrame.TextRange.Text == empty ? noTimeLeft : empty);
                }
                else
                {
                    // Still running ....
                    timeLeft = timeLeft - Constants.RefreshTimeSpan;

                    // Update Shape with new value
                    timerShape.TextFrame.TextRange.Text = timeLeft.ToString("mm\\:ss");

                    // Check treshold and adapt colors
                    // If 25% left => make it orange
                    // If 10% left => make it red
                    double percentageLeft = (100 * timeLeft.TotalMilliseconds) / timerDurationTimeSpan.TotalMilliseconds;
                    if (percentageLeft <= 25 && percentageLeft > 10)
                    {
                        // Weird color conversion here .... so needed to use ToOle
                        timerShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Constants.Less25Percent);
                    }
                    else if (percentageLeft <= 10)
                    {
                        timerShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Constants.Less10Percent);
                    }
                }
            }
        }
    }
}