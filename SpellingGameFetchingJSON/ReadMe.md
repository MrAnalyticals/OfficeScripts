**Excel Office Scipts - using the Fetching JSON functionality - a DEMO**

YouTube video:

Office Scipt .ts  :
OfficeScipt .osts :

![image](https://user-images.githubusercontent.com/47678539/235309969-0eb2e7ec-c25c-47f0-8eb8-439148a67afb.png)

![image](https://user-images.githubusercontent.com/47678539/235310018-588053c5-7b8d-47aa-8681-36995eabd6e8.png)


**Video Audio Script**

Hi, In this video I show you how to build an Excel Office Scripts spelling game that utilizes the fetch function to retrieve JSON data from a REST API. 

Let's now have a look at the demo.


On the spelling sheet we have a table with 4 coloumns. In column b we have 5 words with some of the letters being starred out. we can set the number of stars in the table to the right in col h. You just enter 1, 2 or 3 in there.
 

In column E we have the definition for the words in column b. The definition is retrieved from a rest api using the fetch function in the office script. The way to play this game is to enter your suggested word into column c. 
Column d has a formula in it which will tell you if you are correct or not. 
You can generate a new set of words by clicking the green button.
And, here, we can see that the script has finished and changed each of the words here and created some new definitions. 
Lets try, now, entering some answers. 
And we can see that the formula returned Correct thus confirming that my answers were, indeed, correct.   

Let’s now have a look at the script itself. We can see at the beginning it starts with the async function It returns a promise of data type string or string constructor.  
Next we declare the variables which includes the tables and sheets. 
Next we select 5 words at random from the word list. The word list is contained within the table called hard words. 
We, then, declare the variable difficulty which we are using to determine the number of characters starred out from the words in the sheets. 
Next we populate the sheet with the starred out words. 
Then we enter into the section where we are calling the rest A P I. Here we have a loop. We are iterating through the selected words array. The URL for the rest API is API dot dictionary api dot dev. We are entering the parameter at the end. Then, we are using a try catch to capturing any error responses which can happen time to time especially when dealing with the internet dependent code like a REST A P I. Then we can see, here, we are using the await function, fetching the url, then awaiting the response in Jason format. Then we are using an interface called dictionary a p I response. The interface assists office scripts in handling the Jason schema. It can handle the response to some extent without the interface but this will ensure it gets to all levels within the Jason objects. Next, we are declaring a constant where we are obtaining the definition of the Jason object we require. We have a question mark which means optional. And then we return the definition back to the worksheet. We then loop again. We can see the interface. I generated this by using chat g p t. I pasted in the Jason response object and asked for the Jason schema. 
