
Excel Office Scripts Knights Tour
YouTube video: https://youtu.be/WaLGnBs7QwM
The OfficeScript .ts files can be found here:
The Excel workbook can be found here:
<img width="775" alt="Set Knight Start Pos" src="https://user-images.githubusercontent.com/47678539/175559498-f47ba43a-05f6-4dbb-9775-683b0578fc36.PNG">

Video Script
For those of us who love chess and Excel here is a treat for you! 
A chess themed Excel Office Scripts demo ! 
So, a quick introduction for those of who are new to Office Scripts. Excel Office Scripts is a new macro language for Excel Online. Excel VBA does not work online. Office Scripts was utilised to fill that gap. Office Scripts is a language based on TypeScript but with slight variations because it operates with an Excel object model. 
At the time of producing this video Office Scripts do run on Excel for Desktop for Windows but only in run mode not edit mode.
In this demo I show how Office Scripts can be used to move images around the Excel grid and how it can be used to solve a problem like the Knights Tour. 
The second video explains the code behind the demo.

<img width="960" alt="Wikientry" src="https://user-images.githubusercontent.com/47678539/175559557-d25f1e7f-f141-4605-ba97-bae917071abd.PNG">

So, what is a “Knight’s Tour”. As per the associated Wikipedia page here:
Knight's tour - Wikipedia
https://en.wikipedia.org/wiki/Knight%27s_tour
a knight's tour is a sequence of moves of a knight on a chessboard such that the knight visits every square exactly once.
The knight's tour problem is the mathematical problem of describing and finding a knight's tour. In the code I use Office scripts to attempt to find a Knights tour. 
Because Office Scripts is integrated with Excel it is particular well placed for solving mathematical problems.
The first demo, here, shows how a Knight is moved to a location of your choosing. I enter a location in cell L1 and click the “Set Knight Start Position” button. 
The second demo shows the Knight traversing one random walk. The code uses the random function to choose the direction for the knight. After each move a number is entered into the square to keep track of the movement of the Knight. In cell K4 a message appears stating if a Knight’s Tour was found or not.
For the third demo all we are doing is repeating the 2nd demo, that is repeatedly going on random walks, until it finds a Knights Tour. 
We open the workbook in Windows Desktop in order to facilitate faster running of the code. If the Auto Save option is switched off the code runs, on my machine, around 50 times faster. At this time, Microsoft does not provide the ability to switch auto-save off for the browser version of Excel. There is, presumably, some sort of irrefutable logic behind that decision but it is so complex as to be beyond my reckoning. 
In cell L3 the total number of steps completed on that tour is entered by the code. The steps completed are, also, entered into cell K7. 
And here I stop the code. You could connect the script to a Power Automate Flow and an email action to notify that a Tour was found but I have found that the Flow returns a time out error after 40 minutes or so of the Excel Script action running. If you leave the script running in an open browser window I, also, find that it stops running because Edge has an auto-sleep function for open browser pages.  
Part 2 of this video shows the code description.

