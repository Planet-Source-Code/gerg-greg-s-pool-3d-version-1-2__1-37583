GREG'S POOL 3D – PROTOTYPE VERSION 1.2 		(11.08.2002)

IMPROOVEMENTS IN VERSION 1.1:
-	Collision detection - balls don't fly off into infinity any more (hopefully)
-	Camera control - during aiming the camera is slightly higher than before 
-	More dynamic sound - greater differences of volume between high and low energy collisions

IMPROOVEMENTS IN VERSION 1.2:
-	Enhanced playability with basic European rules for eight-ball pool (provided by Tash)
-	Some general improvements in user interface and game settings
-	Smaller number of spelling mistakes in the Readme file

ACKNOWLEDGMENTS:
-	Version 1.2 would not be created at all if it wasn't for Tash, 
	who was kind enough to provide me with basic rules of European Pool 
	and then volunteered for the function of a guinea-pig.
-	I would also like to thank Ulli for his handful of good ideas.
-	And last, but not least I would also like to thank the whole team behind 
	the GIMP 1.2.3. Most of the “art work” You see in this game was created 
	with the GIMP.


(Now, the old stuff:)

NOTE: English is not my native language, thus You may find many spelling and grammatical mistakes in this text as well as within the code itself. 

DISCLAIMER: Though this programme was tested and caused no problems, I cannot guarantee, that it will run smoothly on Your machine. You use this code on Your own risk.

TO THE POINT - GAME'S DISCRIPTION:
This is a simple game of pool with Direct3D graphics. It is almost complete and performs generally well (at least on my computer – a 650MHz processor with 192MB of RAM and 16MB on a Riva TNT2 graphics card), though it still has some loose-ends (listed later).

The game makes use of basic 3D techniques, like:
-	textured meshes, 
-	vertex and index buffers, 
-	alpha blending, 
-	matrix transformations,
-	billboards 
-	directional lighting

Other built-in features include:
-	2D physics with a collision detection and response mechanism
- 	sprites
-	custom controls
- 	mobile cameras 
- 	dynamic sound
 	
Known (and still unhandled) bugs and loose-ends include:
-	Looks well only on 1024x768 resolution
-	Fails, when the Direct3D device is lost
-	Does not have a help system
-	The error handling mechanism works, but is oversimplified
-	The table could use some details


Game controls:
The main input device is the mouse. For moving the camera press the left or right mouse button (depending on the type of movement you want) and move the mouse.	
There are also few keys that can be used:
Tab	- toggles between available cameras
Space	- launches the cue-ball
F2	- starts a new game
ESC	- exits


FINAL NOTE: I am not planning on further developing this project (maybe except "killing" some bugs). If You like it and have some ideas on how to make it better, feel free to take it and change whatever You like. If you have any questions, comments or remarks feel free to mail me.

Have fun 
Author: Grzegorz Holdys (Wroclaw, Poland)
E-mail: gregor@kn.pl
