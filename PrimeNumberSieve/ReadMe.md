
**Using the Sieve of Eratosthenes to find the Nth Prime numeb using Excel Office Scripts**

YouTube Video URL: https://www.youtube.com/watch?v=R2qX6bInsG0


Office Script: https://github.com/MrAnalyticals/OfficeScripts/blob/main/PrimeNumberSieve/PrimeGenSieveEratosthenesv2.osts



The Sieve of Eratosthenes is an algorithm used to find all prime numbers up to a given value. It works by creating a list of all numbers up to the given value and then repeatedly crossing out (or "marking") all multiples of each prime number as they are discovered, leaving only the prime numbers at the end.

For example, if we want to find all prime numbers up to 30, we would start by creating a list of all the numbers from 2 to 30:
We then start with the first prime number (2) and mark all of its multiples in the list (4, 6, 8, 10, etc.):
Next, we move to the next unmarked number (3) and mark all of its multiples in the list (6, 9, 12, 15, etc.):
We repeat this process with each unmarked number until we reach the end of the list. The remaining unmarked numbers are all prime:
This method is more efficient than simply checking every number up to the given value for primality, especially for larger values of n. The time complexity of the algorithm is approximately O(n log log n), where n is the input value.

In this code I define two functions, sieve Of Eratosthenes and Nth Prime, and a main function that uses these functions to find the nth prime number and output the result to cell B2. The sieve Of Eratosthenes function, as per its name, uses the Sieve of Eratosthenes mathematical algorithm to generate an array of prime numbers up to a given input value n. The Nth Prime function calls sieve Of Eratosthenes function to generate an array of prime numbers up to the square root of n and, then, returns the nth prime number from that array. The main function takes an Excel workbook as input, reads an integer value from cell B1 of the Sheet1 worksheet, then calls Nth Prime function with that input value, and writes the resulting prime number to cell B2. Finally, the main function also logs the value of the prime number to the console.
Let’s see it in action. 
Here I test the script with n for 1, 100, 1000 and finally for 10,000. This method is considerably slower than the method I demoed in my previous video which used the Trial by Division method. 
The script can be found in the GitHub page listed here and in the video’s description. 

![image](https://user-images.githubusercontent.com/47678539/219746439-84b65b98-c37d-40a6-981d-e5d15007cbd7.png)


