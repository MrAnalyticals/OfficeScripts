**"Twinkle twinkle little star. How I wonder what you are" -  Wolfgang Amadeus Mozart**

**An Office Scripts Animation Demo**


![image](https://user-images.githubusercontent.com/47678539/215289777-460cf1d2-22f4-4824-8396-751559a13ad8.png)

**YouTube video:**  

**Office Script:** https://github.com/MrAnalyticals/OfficeScripts/blob/main/StarryNight/StarryNight2.osts

In this video I demo how to create animation in Excel using Office Scripts. As you can see we have twinkling stars as well as a moving UFO image. 
Getting the speed of the UFO and the rate of the twinkling stars was a process of trial and error as each affected the other and also differed depending on which Excel platform was used. The animation ran faster using Excel desktop.
The script can be found on the GitHub repo referenced in the description of this video. 
Let's have a look at the script itself.

We start the script creating the grid variable filling it with empty strings.
Next we use a for loop to populate 150 stars into the grid variable. We use the random function for this. 
We then populate the worksheet with the result of that grid.

Next is the main part of the script that implements the logic of the associated created functions. The variables i and j had to be adjusted by trial and error to adjust the speed of the twinkling stars and the ufo craft whilst , also, testing this on the two Excel platforms, Online and desktop. 
We, also, have a move UFO function which uses the j variable. It is not possible to have two scripts running at the same time or two functions so the two actions, the craft moving and the twinkling have to be done within the one loop. As we can see here.
The replace star with empty function implements the twinkling action. 
Let's have alook at it. 
In the funciton we randomly select cells in the grid and set them to be empty and then back to starred. We use a console dot log command to add delay. Oddly, if I remove this line I get an error in the script telling me the resource has reached capacity! so, I kept it in! 
Finally I have the move UFO function. This uses a shape and I am setting its left and top properties with a parameter created in the previous, governing, for loop.

