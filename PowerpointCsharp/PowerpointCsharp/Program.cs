using System;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;

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

            //Dealing with shapes
            PowerPoint.Shapes shapes = slide.Shapes;  //taking all the shapes collection from my current slide                                        
                                                      //Deleting all shapes
                                                      //delete_shapes(shapes); //calling a new function to delete the shapes from the current slide
            //Add pictures
            addPictures(slide);

            //Add TextBox
            add_textbox(slide);

        }

        private static void add_textbox(_Slide slide)
        {
            PowerPoint.Shape[] shape = new PowerPoint.Shape[10]; //creating a local collection of shapes

            shape[0] = slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                100, 200, 250, 50); //saving the returned shape into a vector of shape (not of shapes, but shape)

            //Ccreating a TextRange - our objet text
            PowerPoint.TextRange textRange;
            textRange = shape[0].TextFrame.TextRange; //assigning our shape to our text range
            textRange.Text = "Here goes the text";
            textRange.Font.Name = "Helvetica";
            textRange.Font.Size = 12;
            textRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue; //for bold text
            textRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue; //or italic text

            //Changing text color
            SolidBrush Brush = new SolidBrush(Color.FromArgb(255, 10, 200, 240));
            textRange.Font.Color.RGB = Brush.Color.ToArgb();
        }

        private static void addPictures(_Slide slide)
        {
            slide.Shapes.AddPicture(@"C:\Users\lucaslemos\Desktop\IMG_example.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
            200, 200, 30, 40); //the first paramether is the local address, the socnd and third I recommend to set up like this
                               //and the you have Left, Top, Widht and Height
        }

        private static void delete_shapes(Shapes shapes)
        {
            if (shapes == null)
                return;
            List<PowerPoint.Shape> listShapes = new List<PowerPoint.Shape>(); //using System.Collections.Generic is necessary
            foreach (PowerPoint.Shape shape in shapes)
            {
                listShapes.Add(shape); //adding each shape on my list of shapes
            }
            foreach (PowerPoint.Shape shape in listShapes)
            {
                shape.Delete(); //deleting one by one
            }
        }
    }
}
