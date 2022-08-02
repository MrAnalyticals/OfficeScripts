
There is no Power Query for Excel Online so how did I import 2.7M rows?
With a little help from Power Automate, Data Factory and Office Scripts of course.
Can Excel OnLine cope with Big Data? Yes, with a little help from Power Automate and Data Factory.
YouTube Video: https://youtu.be/SuFYixYWE2M 
The ImportCSVasParam Office Script is here:  
The Table of Contents Office Script : https://github.com/MrAnalyticals/OfficeScripts/blob/main/ExcelContentsPagev2/TableOfContentsV2.ts  

The learning points from this video are as follows:

•	Azure Data Factory 
o	Split a file into smaller files using the “Max rows per file” Copy-Data setting.
o	Send data to a file using the Copy-data “Additional Column” method. 
o	Trigger a Data Factory pipeline when a file is modified in Azure Data Lake
o	Using variables and parameters in a Pipeline
•	Power Automate
o	Pass data into an Excel Office Script from a Power Automate Flow.
o	Using Azure Blob Actions in Power Automate.
o	Using String functions in Power Automate. 
•	Excel Office Scripts
o	Handling arrays of data with For Loops.
o	Try-Catch Error Handling 

Video Transcript

In a recent video I demonstrated use of the Office Scripts Fetch function to webscrape  csv data from GitHub. I, also, showed the size limitations of that method. In this new video I demo a small project that enables the importation of as many rows, of data, as you like into Excel Online, as long as the Excel file, itself, remains smaller than 100MB. 
My demos show the importation of two datasets one of over 300,000 rows and a larger one of 2.7 M rows.
Excel Online does not have, at this time, a functioning Data Menu. That is, we cannot import data. So, using Power Automate, Azure Data Factory and Excel Office Scripts I will show you how you can bypass that lack of functionality and build your own Data importing solution. 
Microsoft is as we speak building and releasing at various stages importing functionality. 
Because Power Automate’s actions are by their very nature HTTP limited it means when passing data to and from http located data sources it will never be as capable as Azure Data Factory. Data Factory can perform the ETL and integration activities without passing data to and from different http locations. Here are some of the limitations of Power Automate when it comes to handling rows of data. In this project I use Data Lake stored data and Data Factory to do the transforming tasks. Power Automate is used only after the data has been cut down (also known as chunked) to within its size limits. 
Because I am dealing with big data, here, in this case up to 2.7 million rows, looping through so many rows, Power Automate cannot cope with that many rows, even when the flow owner has the most expensive licence. 
