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
    I chose Console App, even though I won't use the console itself. Feel free to choose what you want (and you know about it)
    
2) Install the new library 
    Click with the right button on the project name (the menu on the right) 
    Then go to "Manage NuGet Packages"
    After that, chose the tab "Browse" 
    Search for " ". Install the package
    
3) Add "#using ..." on the top of your "Program.cs" function


Once you have done those steps, you are about to start the presentation.

First, decide if you want to create a new presentation or use a local Powerpoint file as template. 
I recommend to use the a local file if you want to have some slides that will not be touched, for example cover and the "thank for your attention" slide. 


Next info coming...