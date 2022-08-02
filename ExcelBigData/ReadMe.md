
**There is no Power Query for Excel Online so how did I import 2.7M rows?**

With a little help from Power Automate, Data Factory and Office Scripts of course.

Can Excel OnLine cope with Big Data? Yes, with a little help from Power Automate and Data Factory.

YouTube Video: https://youtu.be/SuFYixYWE2M

The ImportCSVasParam Office Script is here: https://github.com/MrAnalyticals/OfficeScripts/blob/main/ExcelBigData/ImportCSVasParam.ts 

The Data Factory Pipeline JSON: https://github.com/MrAnalyticals/OfficeScripts/blob/main/ExcelBigData/Pipe_Chunking.txt 


The Table of Contents Office Script : https://github.com/MrAnalyticals/OfficeScripts/blob/main/ExcelContentsPagev2/TableOfContentsV2.ts  


The learning points from this video are as follows:


•	**Azure Data Factory**

o	Split a file into smaller files using the “Max rows per file” Copy-Data setting.

o	Send data to a file using the Copy-data “Additional Column” method. 

o	Trigger a Data Factory pipeline when a file is modified in Azure Data Lake

o	Using variables and parameters in a Pipeline

•	**Power Automate**

o	Pass data into an Excel Office Script from a Power Automate Flow.

o	Using Azure Blob Actions in Power Automate.

o	Using String functions in Power Automate. 

•	**Excel Office Scripts**

o	Handling arrays of data with For Loops.

o	Try-Catch Error Handling 


Video Transcript

In a recent video I demonstrated use of the Office Scripts Fetch function to webscrape  csv data from GitHub. I, also, showed the size limitations of that method. In this new video I demo a small project that enables the importation of as many rows, of data, as you like into Excel Online, as long as the Excel file, itself, remains smaller than 100MB. 
My demos show the importation of two datasets one of over 300,000 rows and a larger one of 2.7 M rows.
Excel Online does not have, at this time, a functioning Data Menu. That is, we cannot import data. So, using Power Automate, Azure Data Factory and Excel Office Scripts I will show you how you can bypass that lack of functionality and build your own Data importing solution. 
Microsoft is as we speak building and releasing at various stages importing functionality. 
Because Power Automate’s actions are by their very nature HTTP limited it means when passing data to and from http located data sources it will never be as capable as Azure Data Factory. Data Factory can perform the ETL and integration activities without passing data to and from different http locations. Here are some of the limitations of Power Automate when it comes to handling rows of data. In this project I use Data Lake stored data and Data Factory to do the transforming tasks. Power Automate is used only after the data has been cut down (also known as chunked) to within its size limits. 
Because I am dealing with big data, here, in this case up to 2.7 million rows, looping through so many rows, Power Automate cannot cope with that many rows, even when the flow owner has the most expensive licence. 

![image](https://user-images.githubusercontent.com/47678539/182425363-85952108-bebd-4d9f-901a-fada7d2647a7.png)


**Word System Design diagram**
![image](https://user-images.githubusercontent.com/47678539/182425467-236ad117-63bd-4603-899f-6b173aa1d9a8.png)


As shown in the solution diagram the process starts with the file being uploaded to the required Storage Container. This can be done manually through the Azure Portal itself or via the Storage Explorer or indeed by other automated methods. When the file arrives a Data Factory pipeline is triggered which splits that large file into files of 40,000 rows long. Those chunk files are created in a folder called “Chunked”. Once the chunking process has finished a txt file is updated to include a list of the names of the newly created chunk files. This is used to create the location string of those chunk files. Once the txt file has been updated a Power Automate Flow is triggered which then loops through each of the chunk files passing their contents (i.e. their data) into an Office Script. That Office Script outputs the data into a new worksheet for each chunk file. 
You can, if you wish, modify the Office Script to append each chunked input to the end of the previous chunk but given the row limit in Excel worksheet is 1 million rows this will fail for our second upload of 2.7 million rows. You can also, amend the Office Script to overwrite any existing table with the newly inputted chunk data. And, indeed, doing that would, in effect, create a refresh process.
Additionally you can add a DropBox trigger to the Flow so that when you or, indeed, a client, drops a file into a shared DropBox, the file will get transferred into Azure Data Lake. There is a 50MB limit on the Power Automate Drop Box Trigger. 

![image](https://user-images.githubusercontent.com/47678539/182425563-c0f27516-c1cb-4588-af8b-acb84f1775bf.png)


![image](https://user-images.githubusercontent.com/47678539/182425592-c2bedecd-8926-445f-855f-29409c76268f.png)




**The Azure Data Factory**

![image](https://user-images.githubusercontent.com/47678539/182425733-8eab88b1-a23b-4f34-8d73-bfd4477334c9.png)

See Pipe_Chunking.txt file for full JSON version of the pipeline as well as its Activities

**The Power Automate Screenshots**

![image](https://user-images.githubusercontent.com/47678539/182425817-0bfa740a-929a-4f3c-bdbb-c0d2dc367e9a.png)


![image](https://user-images.githubusercontent.com/47678539/182425844-58a510d4-09ab-4a21-bd87-5ab98bd45f23.png)


![image](https://user-images.githubusercontent.com/47678539/182425890-9a659e81-d6d3-4d9e-a672-2c49533fbcf9.png)


![image](https://user-images.githubusercontent.com/47678539/182425919-b236bc58-1119-439f-94d7-99f71c23fbff.png)


![image](https://user-images.githubusercontent.com/47678539/182425955-3fafbf3a-e608-4bff-af14-8ecbf7e72f9e.png)



