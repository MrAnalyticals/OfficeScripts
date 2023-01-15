**File Log File (a “Falafel”) Flow and Office Script**


Keep track of changes to all files in your SharePoint site Library with the help of Power Automate and Excel Office Scripts. 
![image](https://user-images.githubusercontent.com/47678539/212504667-e188a9b3-5895-40b7-848a-ddfb20d71771.png)


**Video Dialog**

Welcome to another Office Scripts video. Hello, I am Emily the Irish Robot host for your video today. So, what is a Falafel? It is a File Log File. So a Falafel is an acronym for a file that keeps a log of other files. 
In this case it is an Excel file. In the Office Script we collect some information about each file. This information is input as parameters obtained from the associated Power Automate Flow.

**Solution Plan**

1.	Triggered by modification or creation of any file in a SPO Library. 
2.	Test for name of file is or is not the Falafel file. 
3.	Run an Office Script using the information obtained from the trigger action, that creates a row in Excel Table for new files and updates for existing ones.

So, lets have a look at the flow. It is triggered by a SharePoint action. Next we have a condition. The file log file itself is called falafel and we don't want to run this flow when that file is modified as this will give rise to an infinite loop. So hence this condition. So, when the condition is false I then run the script. But before that I need to obtain a unique identifier and that is obtained from the item link for the file. Every file has a unique identifier and I have used a formula to obtain the last portion of the item link which is the unique identifier. So lets have a look at the script action. We are obtaining quite a few different values from the trigger. The script is called file logging. It runs against the falafel file itself. The script accepts parameters as follows : the file name with extension, modified date, modified by display name created date by display name , unique identifier, link to the item, the folder path and the version. All of this will be input into the file log file itself. 
So, in the file log file we can see the table that has been formatted, the column headers. and we can see the data that has been imported by the flow. 
Each time a file is modified or created its data is added to this table. So, lets now have a look at the script itself. 
So, we start of with the main function. We have a lot of different parameters. We are declaring the worksheet and the table and we are setting a formula in the cell in the worksheet. This is to obtain the matching record number. We are using the unique identifier as the search term. Once we know the found or matching row number, in the table, we are then able to determine which row we should be updating. For the case of adding a row we, simply, add it to the end of the table. So, going further into the script the first thing we are doing is testing the cell that contains the matching formula to determine if it found a match or not. If it did not find a match it returns the error has n eh and that means we should add a row. That is what were are doing here. And we are adding each of the parameter values as a column value. If it does find a match we, then, obtain the row number which is just the value in the formula cell. We add two because the table starts at row 2. We then update the table row and that finishes the script. 

We Capture in the Excel file the following information:
•	Modified
•	Modified by DisplayName
•	Created
•	Created by DisplayName
•	Identifier
•	Link to Item
•	Filename with extension
•	Folder path
•	Version Number
•	Modified by



