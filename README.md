# Latex4PowerPoint
This is a Latex Add-In for Powerpoint. It enables to add and edit Latex equations or symbols in a Powerpoint slide easily. The Add-In is based on ScintillaNET and supports syntax highlighting, code snippets etc.

With the add-in you can open a dialog, enter something like $y = \sum_{i=0}^n x_i$ and you will get an equation in your slide. This is done by embedding  your Latex code in a template. Then, latex.exe is called in order to create a DVI which is afterwards converted in a PNG file by Ghostscript. The PNG is then inserted in your slide. Your Latex code is stored in  the generated PowerPoint object. Hence, you can edit your code later.

![screenshot](screenshot.jpg)

## Install Add-In

Download one of the release files and start the setup to install the add-in. The add-in should then be loaded automatically when starting PowerPoint the next time. 

Be sure that you install the correct version (x86 or x64).

## Build Add-In

First, you need Visual Studio with C# support. We tested the build using Visual Studio 2019. 

To build the add-in perform the following steps:

* clone repository from GitHub
* open solution file in Visual Studio
* select correct platform (x86 or x64) depending on your PowerPoint version
  * If you are not sure if you have a 32bit or 64bit version of PowerPoint, just start PowerPoint and open "File" -> "Account" -> "About PowerPoint".
* build the add-in
* start PowerPoint and the add-in should be loaded automatically