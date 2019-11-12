using System.Drawing;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointTimer.Core;

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
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    2, 2, 120, 45);

                textBoxCounter.TextFrame.TextRange.Text = Constants.DefaultTimerValue;
                textBoxCounter.Tags.Add(Constants.TimerTag, Constants.TimerTagValue);
                textBoxCounter.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                // Hide it
                textBoxCounter.Visible = MsoTriState.msoFalse;

                var textBoxLauncher = currentSlide.Shapes.AddTextbox(
                   MsoTextOrientation.msoTextOrientationHorizontal,
                   10, 10, 120, 45);

                textBoxLauncher.TextFrame.TextRange.Text = Constants.DefaultTimerValue;
                textBoxLauncher.TextFrame.TextRange.Font.Size = 40;
                textBoxLauncher.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                textBoxLauncher.Tags.Add(Constants.TimerLauncherTag, textBoxCounter.Id.ToString());

                // Add a "bold" animation to triggers the countdown
                currentSlide.TimeLine.MainSequence.AddEffect(textBoxLauncher,
                    MsoAnimEffect.msoAnimEffectBoldFlash, MsoAnimateByLevel.msoAnimateTextByAllLevels,
                    MsoAnimTriggerType.msoAnimTriggerMixed);
            }

        }

        private void RemoveTimers_Click(object sender, RibbonControlEventArgs e)
        {
            // Loop in the slide to remove all timers (launcher and actual timers)
            var app = Globals.ThisAddIn.Application;
            if (app.ActivePresentation.Slides.Count > 0)
            {
                // Get current slide
                Slide currentSlide = app.ActiveWindow.View.Slide;

                currentSlide.Shapes.OfType<Microsoft.Office.Interop.PowerPoint.Shape>()
                    .Where(s => s.Tags[Constants.TimerTag] == Constants.TimerTagValue || s.Tags[Constants.TimerLauncherTag] != "")
                    .ToList()
                    .ForEach(shape => shape.Delete());
            }
        }
    }
}
