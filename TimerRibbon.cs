using System.Drawing;
using System.Linq;
using MOO = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointTimer.Core;
using System.Threading;

namespace PowerPointTimer
{
    public partial class TimerRibbon
    {
        private void AddTimer_Click(object sender, RibbonControlEventArgs e)
        {
            // check if the presentation has slides
            var app = Globals.ThisAddIn.Application;
            if (app.ActivePresentation.Slides.Count > 0)
            {
                // Get current slide
                Slide currentSlide = app.ActiveWindow.View.Slide;

                // Add timer to this slide
                var textBoxCounter = currentSlide.Shapes.AddTextbox(
                    MOO.MsoTextOrientation.msoTextOrientationHorizontal,
                    2, 2, 120, 45);

                textBoxCounter.TextFrame.TextRange.Text = Constants.DefaultTimerValue;
                textBoxCounter.Tags.Add(Constants.TimerCounter, Constants.TimerCounterValue);
                textBoxCounter.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                // Hide it
                textBoxCounter.Visible = MOO.MsoTriState.msoFalse;

                var textBoxLauncher = currentSlide.Shapes.AddTextbox(
                   MOO.MsoTextOrientation.msoTextOrientationHorizontal,
                   10, 10, 120, 45);

                textBoxLauncher.TextFrame.TextRange.Text = Constants.DefaultTimerValue;
                textBoxLauncher.TextFrame.TextRange.Font.Size = 40;
                textBoxLauncher.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                textBoxLauncher.Tags.Add(Constants.TimerLauncher, Constants.TimerLauncherValue);

                // Group both Shapes
                var range = currentSlide.Shapes.Range(new string[] { textBoxLauncher.Name, textBoxCounter.Name });
                var timerShape = range.Group();

                timerShape.Tags.Add(Constants.TimerGroup, Constants.TimerGroupValue);

                // Add a "bold" animation to triggers the countdown
                currentSlide.TimeLine.MainSequence.AddEffect(timerShape.GroupItems.OfType<Shape>().FirstOrDefault<Shape>(s => s.Tags[Constants.TimerLauncher] == Constants.TimerLauncherValue),
                    MsoAnimEffect.msoAnimEffectBoldFlash, MsoAnimateByLevel.msoAnimateLevelNone,
                    MsoAnimTriggerType.msoAnimTriggerMixed);
            }

        }

        private void RemoveTimers_Click(object sender, RibbonControlEventArgs e)
        {
            // Loop in the slide to remove all timers (launcher and actual timers stored in a group)
            var app = Globals.ThisAddIn.Application;
            if (app.ActivePresentation.Slides.Count > 0)
            {
                // Get current slide
                Slide currentSlide = app.ActiveWindow.View.Slide;

                currentSlide.Shapes.OfType<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .Where(s => s.Tags[Constants.TimerGroup] == Constants.TimerGroupValue)
                    .ToList()
                    .ForEach(shape => shape.Delete());
            }
        }
    }
}
