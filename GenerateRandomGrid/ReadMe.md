

**Excel Office Scripts Generate a Random grid of size n**

![image](https://user-images.githubusercontent.com/47678539/214701573-df36ac55-3c6c-49ee-84d3-109b8c36c3c6.png)

**YouTube Video**: https://youtu.be/EvMFUTMeD44

**Video Script**
In this video I demonstrate how to generate a grid of random numbers of size n with the added paramterisation of the range of those random numbers. Excel has a limit of 50 columns for a number range of 1 to 1000. Though you may find a different limit. For extra interest I have added functionality to remove all those numbers that are prime from the grid.
So, let's now have a look at the code. We have a main function where we initiate the routine. You can see we input the parameters for the size of the Grid and specify the lower and upper limits of the input values.

There are two functions, Create random grid and Is prime. Create random grid uses the is prime function. Each time this function is run it clears any existing grid. Then it enters into two For loops. The first iterates through the rows and the second through the columns. It uses the random function to generate the numbers. For each number it tests if it is prime and clears the value from the cell. 

Here, we can see the Is prime function.  This function uses a for loop to check if the number is divisible by any number between 2 and itself. If the number is divisible by any of those numbers, it is not prime and the function returns false.

The Office Script is here: https://github.com/MrAnalyticals/OfficeScripts/blob/main/GenerateRandomGrid/GenerateRandomGrid.osts
