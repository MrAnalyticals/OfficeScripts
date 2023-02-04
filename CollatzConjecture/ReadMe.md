**Number Theory. Introducing the Collatz Conjecture using Excel Office Scripts**



The Collatz Conjecture is a simple mathematical problem that has puzzled mathematicians for many years. The problem states that if you start with any positive integer, and if the number is even, divide it by 2, and if it's odd, multiply it by 3 and add 1, and repeat this process, you will eventually reach the number 1. Despite its simplicity, no one has been able to prove or disprove the conjecture, and it remains one of the most famous unsolved problems in mathematics.
For example if we start with the number 3 we follow seven steps to reach one. 
Step 1.	3 x 3 + 1 = 10
Step 2.	10 / 2 = 5
Step 3.	5 x 3 + 1 = 16
Step 4.	16 / 2 = 8
Step 5.	8 / 2 = 4
Step 6.	4 / 2 = 2
Step 7.	2 / 2 = 1

![image](https://user-images.githubusercontent.com/47678539/216773578-1cc2127d-9929-492d-aa43-dd80071d478a.png)

**The Office Script**: https://github.com/MrAnalyticals/OfficeScripts/blob/main/CollatzConjecture/Collatz.osts

**YouTube Video**:

**Audio Script**

The Collatz Conjecture is a simple mathematical problem that has puzzled mathematicians for many years. Despite its simplicity, no one has been able to prove or disprove the conjecture, and it remains one of the most famous unsolved problems in mathematics.
The problem states that if you start with any positive integer, and if the number is even, divide it by 2, and if it's odd, multiply it by 3 and add 1, and repeat this process, you will eventually reach the number 1.
The conjecture has been verified for every number less than 2 to the power of 68.
So, lets demo the script. I input integers from 1 to 300 as shown here. The input number as well as the count of steps taken to reach one is returned in two columns, Number and Step count respectively. In Excel Online the limit is 250,000 input values before it errors out. 
I have created an x , y scatter chart showing the relationship between each input number and its step count. You can see there are two focal points with a recognizable patter but then rapidly leading to a chaotic distribution. 
You can at this point continue to examine the data for statistical behaviours. 

