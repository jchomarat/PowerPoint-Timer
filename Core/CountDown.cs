using System;
using System.Drawing;
using System.Globalization;
using System.Security.Policy;
using System.Timers;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointTimer.Core
{
    public class CountDown
    {
        string timerDuration = String.Empty;
        TimeSpan timerDurationTimeSpan;
        Timer timer;
        Shape timerShape;
        bool isValid = true;

        readonly string empty = " ";
        readonly string noTimeLeft = "00:00";

        public CountDown(Shape TimerShape)
        {
            if (TimerShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                timerShape = TimerShape;

                // Get duration
                timerDuration = timerShape.TextFrame.TextRange.Text;
                
                // Convert it to TimeSpan
                if (!TimeSpan.TryParseExact(timerDuration, "mm\\:ss", CultureInfo.InvariantCulture, out timerDurationTimeSpan))
                {
                    isValid = false;
                }
            }
            else isValid = false;
        }

        public int UnderlyingShapeId { get { return timerShape.Id; } }

        public bool EligibleToStart { get; set; } = false;

        public void Start()
        {
            if (isValid)
            {
                // Stop animation
                timerShape.AnimationSettings.Animate = Microsoft.Office.Core.MsoTriState.msoFalse;
                
                // Start timer and ignite Count down
                timer = new Timer(Constants.RefreshTimeSpan.TotalMilliseconds);
                timer.Elapsed += OnTimerElapsed;
                timer.Start();
            }
        }

        public void Stop()
        {
            // Restore duration & animation
            timerShape.TextFrame.TextRange.Text = timerDuration;
            timerShape.AnimationSettings.Animate = Microsoft.Office.Core.MsoTriState.msoTrue;
            
            // Kill timer object
            if (timer != null) timer.Dispose();
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
