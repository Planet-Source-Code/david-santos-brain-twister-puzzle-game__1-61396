Brain Twister
=============

Disclaimer
============
The software provided is provided "as is" without any express or implied warranty 
of any kind. Under no circumstances will the author be held liable for any direct,
indirect, incidental, special, exemplary or consequential damages.

This computer software is protected by copyright law and international treaties.

By installing the software on your system, you are agreeing to all the above terms.

The author makes no claims to the originality of the puzzle game.


Intro
=====
Brain Twister is based on a yellow plastic educational toy that tests your 
abstract skills by requiring you to recreate depicted shapes using only
four different shaped pieces by rotating, flipping and aligning them.

I'm not sure if there is one exactly like this in other countries, but there probably is
something similar.


Instructions
============
Click on a shape to select it.
Click shape and drag to move it.
While it's selected, you can click and drag at an empty space to rotate the piece
Right click a piece to flip it.
Click and release in empty space to deselect.
Hold down Spacebar and click drag to move all the objects together at the same time.
Align the pieces so that it resembles the piece shown at the top right.
When the pieces match one of the possible combinations, the puzzle is marked as solved.
Press Ctrl+N or Puzzle>Next to move to the next puzzle
Press Ctrl+B to Puzzle>Back move to the previous puzzle

Solutions are included, but please don't cheat. ;)


ToDo
=============
- create a better looking icon
- standardize puzzle image sizes
- Add graphical interface to move to next puzzle piece?
- Modify flipping and rotation to use special animated popup icons when piece is selected
- Add setting to use old or new rotation method
- Add music?
- Add special debug mode, complete with puzzle creation kit.
- Shadows. Hah.


The works
==============
My goal was to figure out a way to draw the shapes onscreen, be able to rotate them,
and figure out how to snap the pieces together, and have the program work out if
the shapes were aligned in the proper position, no matter how they were rotated.

After a bit of thinking on the problem, my thoughts drifted to shape morphing on Macromedia
Flash.  There were arbitrary points on the shape. I realized I could put points on the shape that would align to other points, and they would match up even if they were flipped or rotated.

I used regions to create the filled shapes.

I drew the shapes in Adobe and plotted the points that I needed to align with points on the 
other pieces.

The program originally read a text file (puz.txt) to read the shape of the puzzles.
The old code is still included.  The new code reads from a binary encoded file.

The original puzzle piece data in puz.txt is stored as follows:

============================
Data	Meaning
============================
PUZ 	Header
n   	total number of pieces in this file (4 in this puzzle)
=============================
1st puzzle piece
=============================
X,Y	Center of rotation of this piece
n	number of points in this piece
p1	X,Y coordinates
p2
.
.
.
.
pn
=============================
Next puzzle piece
=============================
X,Y	Center of rotation of this piece
n	number of points in this piece
p1	X,Y coordinates
p2
.
.
.
.
pn
=============================
etc...
=============================

Pieces.dat is the binary version of puz.txt. I used:

ReadPiecesTXT 
WritePiecesBIN

in Form Load to create my binary data file, then commented those out and used

ReadPiecesBin 


Having created the pieces, I inserted code into the program to show me where each 
point was on the piece. 

The code determines that two points on different pieces align if they are within 8 pixels.
I made the code display a list of which points were aligned when the pieces were placed
in the correct solution for each puzzle. Since there are multiple solutions for each puzzle
I had to write each solution down.

I wrote down these solutions and stored them in the original.pak file which is a comma separated 
values list of puzzles in this format:

<picture filename>,<number of solutions (n)>,<solution 1>,<solution 2>,<...>,<solution N>
<picture filename>,<number of solutions (n)>,<solution 1>,<solution 2>,<...>,<solution N>
<picture filename>,<number of solutions (n)>,<solution 1>,<solution 2>,<...>,<solution N>
<picture filename>,<number of solutions (n)>,<solution 1>,<solution 2>,<...>,<solution N>

The picture filename is the filename of the image that the app will use when displaying the puzzle.
This file can be BMP, GIF or JPG. The corresponding file must exist in the Puzzles folder.

Finally, original.dat stores information on which puzzles have already been solved


Debug mode
==========
Enabling debug mode displays the hotpoints and currently aligned points.
Aligned points are circled with green.
A list of points is shown at the top right in the form

P(N)-P(N)

P is the piece number and N is the point on the piece, so:

0(0)-2(5) 

would mean point 0 on piece 0 is aligned with point 5 on piece 2.


The actual solution strips everything but the numbers so we're left with 

0025

An entire solution would look like:

00250224102010361130203627332834

Because of this point-by-point method of drawing the pieces, and
point by point alignment testing for solutions, the game could be customized use
other shapes, in other configurations.

For feedback, comments and questions, mail me at saintender.geo@yahoo.com

7/6/2005
=============
- Added resource cursors, cursors now change if rotating or translating
- Added a border line to the puzzle pieces
- can add/remove solutions to puzzles and overwrite the solutions database.
- added XP compatible 32-bit icon

7/30/04
=============
- Added notification if user has alread solved the puzzle.
- Added puzzle state loading/saving during start/quit
- Added 45 degree rotation while Shift key is held down
- Added support for choosing Y-based or angled-based rotation method
- Added snap-on sound

7/29/04
=============
- Completed solutions for all puzzles.  I'd like to make a nice user interface,
  before this is ready for a public alpha release.
- Created icon, so we don't have to live with VB's ugly EXE icon
- Changed rotation to use anghle based on position of the cursor around the center of the 
  selected piece, instead of just up and down.

7/28/04
=============
- Found correct? solution for puzzle A-19
- Added solutions for some of the B puzzles 
- Added wooden background.
- Added splash screen

7/22/04
=============
- Fixed piece 3. Still buggy at 45 deg angles but *much* smoother.
  All pieces seem to match ok. 
- Implemented proper snap-on for that smooth look.
- nopped out debug tools (crosshairs, snap circles, snap list) for now


7/1/04
=============
So much for my July 1 deadline on my website...

Thought the solution up yesterday, implemented everything today,
Thought of using regions to perform hit-testing, then realized 
they could be used to draw the polygons as well.

Started drawing with lines, but no way to fill them, then realized 
regions could do that for me.

Cleaned up a lot of code, implemented the puzzle piece object, 
implemented puzzle piece rotation and translation, implemented 
point alignment checking, created initial correct-alignment (solution) maps.

Because of the nature of the puzzle piece and solution files, you could 
easily make all sorts of puzzle games using straight edged geometric figures.

Changed the puzzle data file to binary (actually took up more space) but at 
least no tom dick and harry will accidentally open it up and see plain text.

Changed format of Solution pack to:

<puzzle image filename>,<no. of solutions (n)>,<solution #1>,<solution #2>,...,<solution #n>



Acknowledgements
================
13 PM Enterprises (www.13pment.com) distributor of Brain Twister
J.S. for transcribing the solutions



