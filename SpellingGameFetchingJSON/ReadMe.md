**Excel Office Scripts - using the Fetching JSON functionality - a DEMO**

In this repo and in the video I demonstrate how you can use the Fetch function to retrieve EST API data from an internet endpoint. Here, we ae using the Office Scripts default method: GET. You can, specify different methods if needed. Note the use of a Promise as a return value for the Main function and the .json method together with the interface to parse the response.


YouTube video: https://youtu.be/vFJIJH0ZNlA


Office Script .ts  :
https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/SpellingGameFetchingJSON/SpellingTest4.ts


OfficeScript .osts : https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/SpellingGameFetchingJSON/SpellingTest4.osts


![image](https://user-images.githubusercontent.com/47678539/235309969-0eb2e7ec-c25c-47f0-8eb8-439148a67afb.png)


![image](https://user-images.githubusercontent.com/47678539/235310018-588053c5-7b8d-47aa-8681-36995eabd6e8.png)


**Video Audio Script**

Hi, In this video I show you how to build an Excel Office Scripts spelling game that utilizes the fetch function to retrieve JSON data from a REST API. 

Let's now have a look at the demo.


On the spelling sheet we have a table with 4 coloumns. In column b we have 5 words with some of the letters being starred out. we can set the number of stars in the table to the right in column h. You just enter 1, 2 or 3 in there.
 

In column E we have the definition for the words in column b. The definition is retrieved from a rest api using the fetch function in the office script. The way to play this game is to enter your suggested word into column c. 
Column d has a formula in it which will tell you if you are correct or not. 
You can generate a new set of words by clicking the green button.
And, here, we can see that the script has finished and changed each of the words here and created some new definitions. 
Lets try, now, entering some answers. 
And we can see that the formula returned Correct thus confirming that my answers were, indeed, correct.   

Let’s now have a look at the script itself. We can see at the beginning it starts with the async function It returns a promise of data type string or string constructor.  Next we declare the variables which includes the tables and sheets. 
Next we select 5 words at random from the word list. The word list is contained within the table called hard words. 
We, then, declare the variable difficulty which we are using to determine the number of characters starred out from the words in the sheets.  Next we populate the sheet with the starred out words. 

Then we enter into the section where we are calling the REST API. Here we have a loop. We are iterating through the selected words array. The URL for the REST API is API.dictionaryapi.dev. We are entering the parameter at the end of this. Then, we are using a Try...Catch pattern to capturing any error responses which can happen from time to time especially when dealing with  internet dependent code like a REST API. Then, we can see, here, we are using the await function, fetching the url, then awaiting the response in JSON format. Then we are using an interface called dictionary API response. The interface assists Office Scripts in handling the JSON schema. It can handle the response, to some extent, without the interface but this will ensure it gets to all levels within the retuned JSON object. Next, we are declaring a constant where we are obtaining the definition of the JSON object we require. We have a question mark which means optional. And then we return the definition back to the worksheet. We then loop again. Next, we can see the interface. I generated this by using ChatGPT.
I pasted in the JSON response object and asked for the JSON schema. Note: it took thee attempts to povide the coect schema. So, it was a matter of trial and eo there.
