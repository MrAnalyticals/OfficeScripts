Learn how to create Classes with Excel Office Scripts. My new demo shows how to use a class to generate any shape in Excel adding new properties and methods. 
Please Like, Share and of course Subscribe!  

YouTube Video: https://youtu.be/1Xze3nlqu2c 

This demo shows how to use a class to generate any shape in Excel adding new properties and methods to it. 
A new property called owner is being declared. We create the getter and the setter for that property. Get owner and Set owner. 

setPosition and setSize are two new methods.

This class is using a feature called inheritance. That is, I am using an existing object called shape and using some of its attributes. That is, some of its properties and some of its methods but not all. 

In the main function there are 4 input parameters. The workbook itself, the set text, for the text that displays in the shape, set shape, the shape type, and owner name values. The user inputs these last 3 values using the default input form. 

There is, next, a try catch error handler. If a shape is wrongly named, by the user, the code responds with an error in the console and sets cell a1 with that same error message. 

Later we have a getshapetype function. This function acts a helper lookup function looking up the user input shape type and returning an enum which is a predefined Excel constance. This can't be built using strings so this helper function was required. 

In the second part of the video I have provided a more advanced version of the code creating a 3rd new method for the class. 
The method aligns the newly created shape next to the right most existing shape. 
