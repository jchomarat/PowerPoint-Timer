using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointTimer.Core;

namespace PowerPointTimer
{
    /// <summary>
    /// API PowerPoint VSTO:
    /// https://docs.microsoft.com/en-US/office/vba/api/overview/powerpoint/object-model
    /// </summary>
    public partial class ThisAddIn
    {
        List<CountDown> runningCountDowns = new List<CountDown>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.SlideShowEnd += OnSlideShowEnd;
            Application.SlideShowNextSlide += OnSlideShowNextSlide;
            Application.SlideShowNextClick += OnSlideShowNextClick;
            Application.SlideShowOnPrevious += OnSlideShowOnPrevious;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }


        /// <summary>
        /// Called each time a slide is loaded. Effect will tell if an effect is eligible on the next click,
        /// if so, mark the corresponding timer as "elligible" for start on the next click
        /// </summary>
        /// <param name="SlideShowWindow"></param>
        /// <param name="Effect"></param>
        void OnSlideShowNextClick(SlideShowWindow SlideShowWindow, Effect Effect)
        {
            // Detect which CountDown will start at next click
            if (Effect != null)
            {
                if (!string.IsNullOrEmpty(Effect.Shape.Tags[Constants.TimerLauncherTag]))
                {
                    // For initialized countdowns, make them ready to rumble for next click
                    runningCountDowns
                        .FirstOrDefault<CountDown>(cd => cd.UnderlyingShapeId == int.Parse(Effect.Shape.Tags[Constants.TimerLauncherTag]))
                        ?.Init(Effect.Shape);
                }
            }
            else
            {
                runningCountDowns.FirstOrDefault<CountDown>(cd => cd.Status == CountDownStatusEnum.Ready)
                    ?.Start();
            }
        }

        void OnSlideShowOnPrevious(SlideShowWindow SlideShowWindow)
        {
            if (runningCountDowns.Count > 0)
            {
                runningCountDowns.ForEach(cd => cd.Stop());
            }
        }

        /// <summary>
        /// Called when the slide show ends. This handler kills all countdowns
        /// </summary>
        /// <param name="Presentation">The current presentation</param>
        void OnSlideShowEnd(Presentation Presentation)
        {
            disposeCountDowns();
        }

        /// <summary>
        /// Called when a slide is loaded, in a slide show mode. Kills existing countdown, and start thoses on the
        /// loaded slide
        /// </summary>
        /// <param name="SlideShowWindow">The current slide show window</param>
        void OnSlideShowNextSlide(SlideShowWindow SlideShowWindow)
        {
            // If timers are running, kill them before starting them on the new slide
            disposeCountDowns();

            // If a timer is on the slide, init it!
            var currentSlide = SlideShowWindow.View.Slide;
            runningCountDowns = getCountDowns(currentSlide)
                .Select(shape => new CountDown(shape))
                .ToList();
        }

        /// <summary>
        /// stop all countdowns registered into the runningCountDowns list.
        /// </summary>
        void disposeCountDowns()
        {
            if (runningCountDowns.Count > 0)
            {
                runningCountDowns.ForEach(cd => cd.Stop());
                runningCountDowns.Clear();
            }
        }

        /// <summary>
        /// Returns all Shape of a given slide that are "CountDown"
        /// </summary>
        /// <param name="slide">the slide to look into!</param>
        /// <returns></returns>
        IEnumerable<Shape> getCountDowns(Slide slide)
        {
            return slide.Shapes.OfType<Shape>().Where(s =>
            {
                return s.Tags[Constants.TimerTag] == Constants.TimerTagValue;
            });
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
