# Automated-Powerpoint-with-CSharp

Creating a Powerpoint Presentation with C# - Visual Studio 2017 Community

Let suppose you are NOT trying to make manually a Powerpoint Presentation. Instead, what you want is to create it programatically, because your application is supposed to take the Data from a database and create a PowerPoint presentation with the data. 

So, these are the topics I want to talk about (and show the code):
- Creating a PowerPoint Presentation from VS 2017 (open, create new and save/save as)
- Add pictures, arrows, tables and other things to it.

For that, I used:
- Visual Studio 2017 Community
- Office 2016 (Office 365 is also compatible, no worry)
- Basic C# knoledgement (if you don't have, you can easily get the idea from this tutorial).

So, let's get started!

-------------------------------------------------------------------------------------------------------------------------------------

1) Create a new Project on VS 2017 
    I chose Console App (.NET Framework), even though I won't use the console itself. Feel free to choose what you want (and you know about it). I choose .NET Framework 4.6.1
    Another thing you will need to do is to add a reference to a "office.dll Version:15.0.0.0". For that, download the file:
    https://www.dllme.com/dll/files/office_dll.html
    Then click with the right button on the References and choose "add a reference". Go to browse and choose the Office.dll inside the        zip file you've just downloaded.
    ATTENTION: this will avoid problems like "msotristate is defined as an assembly that is not referenced". (I spent two hours on it)
    
2) Install the new library 
    Click with the right button on the project name (the menu on the right) 
    Then go to "Manage NuGet Packages"
    After that, chose the tab "Browse" 
    Search for " ". Install the package
    
3) Add "using Microsoft.Office.Interop.PowerPoint;" on the top of your "Program.cs" function

4) Add a reference to Microsoft Office 15.0 Object Library on COM tab
    For that, download the reference here in the .dll format:
    https://www.dllme.com/dll/files/microsoft_office_interop_powerpoint_dll.html
    Click with right button on the "Dependencies" on the right menu. Go to add reference.
    Go to browse and choose the .dll file you have downloaded. Then press ok. 
    Now you need to close your project and open again.
    
    
Once you have done those steps, you are about to start the presentation.

First, decide if you want to create a new presentation or use a local Powerpoint file as template. 
I recommend to use the a local file if you want to have some slides that will not be touched, for example cover and the "thank for your attention" slide. 


---------------------------------------------------------------------------------------------------------------------------------------

#Opening a existing PowerPoint

            //Creating an Application
            Application myApplication = new Application();
            //Creating a Presentation - opening a existing PowerPoint
            Presentation myPresentation = myApplication.Presentations.Open(@"C:\Users\lucaslemos\Desktop\Github\PowerPoint-                         CSharp\tutorial_slide.pptx");
            
 You will probably many times the following sentence "Microsoft.Office.Interop.PowerPoint". To make your code look better,
 you can add on the top:
    
    using PowerPoint = Microsoft.Office.Interop.PowerPoint;
 
 So you can spare time and space. It is just a shortcut.
 
 ------------------------------------------------------------------------------------------------------------------------------------
 #Creating a new slide and shapes
 Once you have the "PowerPoint" shortcut (see just above), you can create a variable slide and a variable slides.
 
    PowerPoint.Slides slides; //will be used as the whole collection of my presentation
    PowerPoint._Slide slide; //will be used as my current slide being edited
    
 The first one is a like a vector that contains all the slides on the presentation. The second one I use as my current slide that I am editing. 
 
    slides = myPresentation.Slides; // (big S)
    PowerPoint.CustomLayout customLayout = myPresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
    slide = slides.AddSlide(2, customLayout); //Creating a new slide, which will be second slide of my presentation
                                              //CustomLayout is a mandatory input
 
 Something you need to understand is: everything inside a slide is called shapes. A shape can be a picture, a table, a textbox, and so on...
 When you use currentSlide.Shapes(1), you are trying to get a shape from a bunch of shapes (Shapes). It is complicated, because you may have NO control on the order. Shapes is a vector and you control by index. 
 That is why a prefere to delete all shapes from a new slide (it ALWAYS comes with at least one shape). I like to create a vector of shapes, so I can define, for example, that my shapes(1) is a picture shapes(2) is a textbox. Because if you change them, even though they're both Shape, they have many different attributes. 
 
             //Dealing with shapes
            PowerPoint.Shapes shapes = slide.Shapes;  //taking all the shapes collection from my current slide                                        
            //Deleting all shapes
            delete_shapes(shapes); //calling a new function to delete the shapes from the current slide
            
 And the function delete_shapes would be:
 
             private static void delete_shapes(Shapes shapes)
        {
            if (shapes == null)
                return;
            List <PowerPoint.Shape> listShapes = new List<PowerPoint.Shape>(); //using System.Collections.Generic is necessary
            foreach(PowerPoint.Shape shape in shapes)
            {
                listShapes.Add(shape); //adding each shape on my list of shapes
            }
            foreach(PowerPoint.Shape shape in listShapes)
            {
                shape.Delete(); //deleting one by one
            }
        }
        
 
 ------------------------------------------------------------------------------------------------------------------------------------
 #Add Pictures, Textboxes, Tables, Arrows, Icons,  
 
 PICTURES
 
               slide.Shapes.AddPicture(@"C:\Users\lucaslemos\Desktop\IMG_example.jpg", Microsoft.Office.Core.MsoTriState.msoFalse,                          Microsoft.Office.Core.MsoTriState.msoTrue, 
                200, 200, 30, 40); //the first paramether is the local address, the socnd and third I recommend to set up like this
                                  //and the you have Left, Top, Widht and Height
LocalAddress = @" [insert here the address]  ";
LinkToFile = Microsoft.Office.Core.MsoTriState.msoFalse, if you do not want the link between the picture file and your powerpoint
SaveWithDocument = Microsoft.Office.Core.MsoTriState.msoTrue, if you want to save with your presentation
Left = from the left board of the slide to the point you want, in pixels
Top = from the top of the slide to the point you want, in pixels
Widht = the width of your image, in pixels
Height = the width of your image, in pixels

I do not know why, but for some picturew you cannot scale them, maybe due to format. For .jpg in the example the scalingg did not work, but for .png it does work. 


TEXTBOXES

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
        
  So here you can see that we have the addTextBox method, that is pretty much like the AddPicture (both have
  left, top, width, height attributes. But for the AddTextbox, the first attirbute is the text orientation (there are
  horizontal, vertical, downward, Mixed, Upward, and so on...
  
  About the color changing, i found it really difficult. The solution I got is: create a SolidBrush and define its color.
  
  .ToArgb(a, b, c, d) are: 
  a = from 0 to 255 -> color transparency
  b =  from 0 to 255 -> blue
  c =  from 0 to 255 -> green
  d =  from 0 to 255 -> red
  
  It should be Red Green Blue, but for some reason is alpha(transaparency), blue, green, red. 
So, if you have the RGB from the color you want, just put alpha as 255 and then the blue component, followed by the green and then
the red component. The order is very important. 
  
