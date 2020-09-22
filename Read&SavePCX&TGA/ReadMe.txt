In these zip-file is a project with 4 cls-files that read and save pcx and tga-Files
whithout any dll. You must put this classes and the cbitmap-bas-file to your
project and your project can read pcx an tga pictures.

Example:
Dim a as New Loadtga
a.LoadTGA "c:\test.tga"
a.Drawtga PictureBox1

If a. Autoscale = True then the picturebox become automaticly the dimensions of the picture
You can scale the pictures very easy (see example-code.

with istga and ispcx you can see is the picture really a pcx or tga-picture.


You can read:
1 Bit-pcx
4 Bit-pcx
8 Bit-pcx
24 Bit-pcx
8 Bit-tga
16 Bit-tga
24 Bit-tga
32 Bit-tga
with and without RLE-Encoding

You can save:
24 Bit-pcx
24 Bit-tga
with and without RLE-Encoding
The code is complete. So you can see anything.

If you find bugs or you change something, please send me a e-mail, so
i can also change my cls-files.



Also is in this zip-file a project that show you can play fli or flc-files
in any picture-box, form and some other objects. This is my secondrealize
of this project!!! Rename Fliplayer.dl_ to Fliplayer.dll and register it. It read flic files in memory!! It read all fli and flc-files.
I create this to read small flic-files (smaller than 1 MB). But if you have anough memory
you can play them (of course, i have read files with 16 MB on my new PC an it works).
You can play flic-files from res-files and transparent (see Sample Transparent.vbp).


I make this project with the help of the book "Das neue Handbuch der Grafikformate" from Klaus Holtorf
and some help from the www.

Please vote for me.

If you have questions or suggestions please mail to:
alfred.koppold@freenet.de