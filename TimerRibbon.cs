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
                var textBox = currentSlide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    2, 2, 120, 45);

                textBox.TextFrame.TextRange.Text = Constants.DefaultTimerValue;
                textBox.TextFrame.TextRange.Font.Size = 40;
                textBox.Tags.Add(Constants.TimerTag, Constants.TimerTagValue);
                
                // Add animation so that the ContDown is activated upon click (like next slide)
                textBox.AnimationSettings.Animate = MsoTriState.msoTrue;
                textBox.AnimationSettings.TextLevelEffect = PpTextLevelEffect.ppAnimateByAllLevels;
            }
            
        }
    }
}
