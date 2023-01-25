**Excel Office Scripts Random Sentence Generator**


![image](https://user-images.githubusercontent.com/47678539/214718769-f10a6765-f6af-4bb8-930c-652bb6f809fa.png)

**YouTube Video**: https://youtu.be/7gHJ1YrC6JA

 
**Office Script**: https://github.com/MrAnalyticals/OfficeScripts/blob/main/GeneratePseudoSentences/GeneratePseudoSentences.osts


In this Script I define two helper functions: generateString and generateSentence.
The generateString function takes in a length and creates a random string of that length, with a special case for length of 1.
The generateSentence function creates a sentence by calling the generateString function 15 times, with the number of vowels added to the end of each generated word depending on the length of the word, and then it returns the sentence.
Finally, it creates a variable SentenceResult to store the result of the generateSentence function as a string. It adds that result as a new row to the formatted Excel table.
