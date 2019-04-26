using System;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerpointCsharp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creating an Application
            Application myApplication = new Application();
            //Creating a Presentation - opening a existing PowerPoint      you need to put the @
            Presentation myPresentation = myApplication.Presentations.Open(@"C:\Users\lucaslemos\Desktop\Github\PowerPoint-CSharp\tutorial_slide.pptx");
            //Dealing with slides
            PowerPoint.Slides slides; //will be used as the whole collection of my presentation
            PowerPoint._Slide slide; //will be used as my current slide being edited
            slides = myPresentation.Slides; // (big S)
            PowerPoint.CustomLayout customLayout = myPresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            slide = slides.AddSlide(2, customLayout); //Creating a new slide, which will be second slide of my presentation
                                                      //CustomLayout is a mandatory input


        }
    }
}
