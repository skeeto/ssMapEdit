
*************************************************
               synSpace Map Editor
         Originally created by Mosquito

   This program was written to aid in the
 creation of maps for Synthetic-Reality's game
 called synSpace.
   (see http://www.synthetic-reality.com)

   This is the remake of the originally map
 editor which was created a few years before. It
 was crap.

   If you want to contact Mosquito, the e-mail
 address is mosquito@adelphia.net

        *********************************

                  Terms of Use

   You are free to edit this code and compile
 it in any way you like as long as you meet the
 following criteria:

   1) You must keep the entire code open and
      you must distribute the code with the
      compiled program.

   2) The program must be free. You may not
      ever charge for the use of the program.

*************************************************


Information on the included files:

*readme.txt
	This file!

*mapdefs.dat
	This file helps keep the program up to date. If Dan ever changes a synSpace limit (such as allowing 200 barriers) you can fix the editor by changing the _MAX value. Also in this file, you can set the location of your Arcadia installation. You can also change the number of Undo's avaliable in the editor by chaning the UNDO_MAX property.

*defaults.dat
	This file is really a map file. It is loaded when the program begins and when you create a new map (File>New). The best way to change it is to create the map you want to be defaulted and save the map to defaults.dat.

*barxpar.day
	This contians the xpar values that are provided by the map0.ini map. If Dan creates a new xpar, you can fix it here in this file. This file does not allow comment lines.

*pupstyle.dat
	This is just like the xpar file. It keeps track of all the powerup styles. If Dan creates a new powerup, you can add it to the editor just by editing the file. This file does not allow comment lines.

